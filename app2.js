(() => {
    "use strict";
  
    const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
    const DRIVE_API = "https://www.googleapis.com/drive/v3/files";
    const SELF_CLIENT = "HSC London(Self)";
    const state = { idTokenPayload:null, accessToken:null, workbookId:null, workbookName:null, inventory:[], transactions:[], clients:[] };
    let pendingMovement = null;
    const $ = id => document.getElementById(id);
  
    document.addEventListener("DOMContentLoaded", () => { bindEvents(); setDefaultTimestamp(); waitForGoogle(); });
  
    function bindEvents(){
      $("grant-access").addEventListener("click", requestSheetAccess);
      $("sign-out").addEventListener("click", signOut);
      $("refresh").addEventListener("click", loadLogger);
      $("movement-form").addEventListener("submit", reviewMovement);
      $("movement").addEventListener("change", handleMovementChange);
      $("asset").addEventListener("change", updatePreview);
      $("client").addEventListener("change", updatePreview);
      $("quantity").addEventListener("input", updatePreview);
      $("timestamp").addEventListener("input", updatePreview);
      $("cancel-confirm").addEventListener("click", closeConfirm);
      $("approve-confirm").addEventListener("click", approveMovement);
      document.addEventListener("click", e => { if(e.target.matches("[data-close-confirm]")) closeConfirm(); });
    }
  
    function waitForGoogle(){
      let checks=0; const timer=setInterval(()=>{ checks++; if(window.google?.accounts?.id && window.google?.accounts?.oauth2){clearInterval(timer);initializeGoogle();} if(checks>=150){clearInterval(timer);setAuthStatus("Google services could not be loaded. Check your internet connection.",true);} },100);
    }
  
    function initializeGoogle(){
      if(!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")){setAuthStatus("Add your existing Google Web Client ID to config.js.",true);return;}
      google.accounts.id.initialize({client_id:CONFIG.GOOGLE_CLIENT_ID,callback:handleCredentialResponse,auto_select:true,cancel_on_tap_outside:false});
      google.accounts.id.renderButton($("google-signin-button"),{theme:"outline",size:"large",text:"signin_with",shape:"rectangular",width:280});
      google.accounts.id.prompt();
    }
  
    function handleCredentialResponse(response){
      try{
        state.idTokenPayload=decodeJwtPayload(response.credential);
        $("user-name").textContent=state.idTokenPayload.name||"Google user";
        $("user-email").textContent=state.idTokenPayload.email||"";
        if(state.idTokenPayload.picture){$("user-photo").src=state.idTokenPayload.picture;$("user-photo").classList.remove("hidden");}
        $("google-signin-button").classList.add("hidden"); $("sign-out").classList.remove("hidden"); $("grant-access").classList.remove("hidden");
        setAuthStatus("Signed in. Click the button below to connect to Google Sheets.");
      }catch(e){console.error(e);setAuthStatus("Google sign-in response could not be read.",true);}
    }
  
    function requestSheetAccess(){
      if(!state.idTokenPayload){setAuthStatus("Sign in with Google first.",true);return;}
      const tokenClient=google.accounts.oauth2.initTokenClient({client_id:CONFIG.GOOGLE_CLIENT_ID,scope:CONFIG.OAUTH_SCOPES,callback:async tokenResponse=>{
        if(tokenResponse.error){setAuthStatus(`Google authorization failed: ${tokenResponse.error}`,true);return;}
        state.accessToken=tokenResponse.access_token; $("grant-access").classList.add("hidden"); $("login-card").classList.add("hidden"); $("logger").classList.remove("hidden"); await loadLogger();
      }});
      tokenClient.requestAccessToken({prompt:"consent",login_hint:state.idTokenPayload.email||undefined});
    }
  
    async function loadLogger(){
      if(!state.accessToken)return;
      setSyncStatus("Syncing with Google Sheets...");
      try{
        const workbook=await findWorkbook(CONFIG.WORKBOOK_NAME); if(!workbook)throw new Error(`Workbook "${CONFIG.WORKBOOK_NAME}" was not found in your Google Drive.`);
        state.workbookId=workbook.id; state.workbookName=workbook.name; $("workbook-name").textContent=workbook.name;
        const metadata=await sheetsGet(`/${encodeURIComponent(state.workbookId)}`);
        let titles=(metadata.sheets||[]).map(s=>s.properties.title); const missing=[];
        if(!titles.includes(CONFIG.INVENTORY_SHEET_NAME))missing.push(CONFIG.INVENTORY_SHEET_NAME);
        if(!titles.includes(CONFIG.TRANSACTIONS_SHEET_NAME))missing.push(CONFIG.TRANSACTIONS_SHEET_NAME);
        if(missing.length)await createSheets(missing);
        const [inventoryRows,transactionRows,mainRows]=await Promise.all([getValues(CONFIG.INVENTORY_SHEET_NAME),getValues(CONFIG.TRANSACTIONS_SHEET_NAME),getValues(CONFIG.MAIN_SHEET_NAME)]);
        state.inventory=parseInventory(inventoryRows); state.transactions=parseTransactions(transactionRows); state.clients=parseClients(mainRows);
        renderInputs(); renderAudit(); setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
      }catch(e){console.error(e);setSyncStatus(e.message||"Unable to load spreadsheet.",true);}
    }
  
    async function findWorkbook(name){
      const query=[`name = '${escapeDriveQuery(name)}'`,`mimeType = 'application/vnd.google-apps.spreadsheet'`,`trashed = false`].join(" and ");
      const data=await fetchJson(`${DRIVE_API}?q=${encodeURIComponent(query)}&pageSize=10&fields=files(id,name,mimeType,modifiedTime,webViewLink)`,{headers:authHeaders()});
      return data.files?.[0]||null;
    }
  
    async function createSheets(names){
      await sheetsPost(`/${encodeURIComponent(state.workbookId)}:batchUpdate`,{requests:names.map(title=>({addSheet:{properties:{title}}}))});
      if(names.includes(CONFIG.INVENTORY_SHEET_NAME))await updateValues(CONFIG.INVENTORY_SHEET_NAME,[["Asset","Balance"]]);
      if(names.includes(CONFIG.TRANSACTIONS_SHEET_NAME))await updateValues(CONFIG.TRANSACTIONS_SHEET_NAME,[["Timestamp","Client","Movement","Asset","Quantity","User"]]);
    }
  
    async function getValues(sheetName){const data=await sheetsGet(`/${encodeURIComponent(state.workbookId)}/values/${encodeURIComponent(quoteSheetName(sheetName)+"!A:AE")}`);return data.values||[];}
    async function updateValues(sheetName,rows){const range=`${quoteSheetName(sheetName)}!A1`;return sheetsPut(`/${encodeURIComponent(state.workbookId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:rows});}
  
    function parseInventory(rows){
      if(!rows.length)return[]; const header=rows[0].map(normalizeHeader); const assetIdx=findColumn(header,["asset","asset name","item","type"]); const balanceIdx=findColumn(header,["balance","current balance","stock","quantity"]); if(assetIdx<0)return[];
      return rows.slice(1).map((row,index)=>({rowNumber:index+2,asset:String(row[assetIdx]??"").trim(),balance:balanceIdx>=0?numericValue(row[balanceIdx]):0,assetColumn:assetIdx+1,balanceColumn:balanceIdx>=0?balanceIdx+1:2})).filter(x=>x.asset);
    }
  
    function parseTransactions(rows){
      if(!rows.length)return[]; const header=rows[0].map(normalizeHeader); const idx={timestamp:findColumn(header,["timestamp","date","datetime"]),client:findColumn(header,["client","client name"]),movement:findColumn(header,["movement","type","direction"]),asset:findColumn(header,["asset","asset name","item"]),quantity:findColumn(header,["quantity","qty"]),user:findColumn(header,["user","entered by","email"])};
      return rows.slice(1).map(row=>({timestamp:idx.timestamp>=0?row[idx.timestamp]??"":"",client:idx.client>=0?row[idx.client]??"":"",movement:idx.movement>=0?row[idx.movement]??"":"",asset:idx.asset>=0?row[idx.asset]??"":"",quantity:idx.quantity>=0?numericValue(row[idx.quantity]):0,user:idx.user>=0?row[idx.user]??"":""})).filter(x=>x.asset||x.client);
    }
  
    function findMainHeaderRow(rows){return rows.findIndex(row=>{const h=row.map(normalizeHeader);return (h.includes("cleint")||h.includes("client")||h.includes("collection client"))&&(h.includes("load type")||h.includes("loadtype"))&&(h.includes("planned arrival")||h.includes("plannedarrival"));});}
    function parseClients(rows){
      const headerRow=findMainHeaderRow(rows); if(headerRow<0)return[]; const headers=rows[headerRow].map(normalizeHeader); const idx=findColumn(headers,["cleint","client","collection client"]); if(idx<0)return[];
      return [...new Set(rows.slice(headerRow+1).map(r=>String(r[idx]??"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    }
  
    function renderInputs(){
      $("asset").innerHTML=state.inventory.length?`<option value="">Select asset</option>${state.inventory.map(x=>`<option value="${escapeAttr(x.asset)}">${escapeHtml(x.asset)}</option>`).join("")}`:`<option value="">No assets configured</option>`;
      const clients=state.clients.filter(c=>c!==SELF_CLIENT); $("client").innerHTML=clients.length?`<option value="">Select client</option>${clients.map(c=>`<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`:`<option value="">No clients found</option>`;
      handleMovementChange(); updatePreview();
    }
  
    function renderAudit(){
      const rows=state.transactions.slice(-30).reverse(); $("audit-count").textContent=rows.length.toLocaleString(); $("transactions-body").innerHTML=rows.length?rows.map(item=>`<tr><td>${escapeHtml(formatTimestamp(item.timestamp))}</td><td><strong>${escapeHtml(item.client||"")}</strong></td><td><span class="movement-tag ${movementClass(item.movement)}">${escapeHtml(item.movement||"")}</span></td><td>${escapeHtml(item.asset||"")}</td><td class="num">${formatNumber(item.quantity)}</td><td>${escapeHtml(item.user||"")}</td></tr>`).join(""):emptyRow(6,"No asset movements recorded yet.");
    }
  
    function handleMovementChange() {
        const isDiscard = $("movement").value === "DISCARD";
        const client = $("client");
    
        if (isDiscard) {
            // Automatically assign discarded assets to HSC London(Self)
            client.value = SELF_CLIENT;
            client.disabled = true;
    
            $("client-help").textContent =
                "Discarded assets are automatically recorded against HSC London(Self).";
        } else {
            client.disabled = false;
    
            // Clear the automatic self-client selection when changing
            // back to Received/Sent.
            if (client.value === SELF_CLIENT) {
                client.value = "";
            }
    
            $("client-help").textContent = SELF_CLIENT;
        }
    
        updatePreview();
    }

    
    function reviewMovement(event){
      event.preventDefault();
      const data=readForm(); const error=validateMovement(data); if(error){setMovementStatus(error,true);return;}
      const item=state.inventory.find(x=>x.asset.toLowerCase()===data.asset.toLowerCase()); if(!item){setMovementStatus("That asset is not present in Inventory.",true);return;}
      if(data.movement!=="RECEIVED"&&data.quantity>item.balance){setMovementStatus(`Cannot remove ${data.quantity}. Current ${item.asset} balance is ${formatNumber(item.balance)}.`,true);return;}
      pendingMovement={...data,item,newBalance:data.movement==="RECEIVED"?item.balance+data.quantity:item.balance-data.quantity};
      $("confirm-movement").textContent=movementLabel(data.movement);$("confirm-client").textContent=data.client;$("confirm-asset").textContent=data.asset;$("confirm-quantity").textContent=formatNumber(data.quantity);$("confirm-time").textContent=data.timestamp;$("confirm-balance").textContent=`${formatNumber(item.balance)} → ${formatNumber(pendingMovement.newBalance)}`;
      const warning=$("confirm-warning"); warning.classList.toggle("hidden",data.movement!=="DISCARD"); if(data.movement==="DISCARD")warning.textContent="Discard is permanent in the warehouse balance. Please make sure the quantity is correct.";
      $("confirm-modal").classList.remove("hidden"); setTimeout(()=>$("approve-confirm").focus(),50);
    }
  
    async function approveMovement(){
      if(!pendingMovement)return; const data=pendingMovement; $("approve-confirm").disabled=true; $("cancel-confirm").disabled=true; setMovementStatus("Recording movement...");
      try{
        const user=state.idTokenPayload?.email||state.idTokenPayload?.name||"Google user"; const transactionRange=`${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A:F`;
        await sheetsPost(`/${encodeURIComponent(state.workbookId)}/values/${encodeURIComponent(transactionRange)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,{values:[[new Date().toISOString(),data.client,data.movement,data.asset,data.quantity,user]]});
        const balanceCell=columnLetter(data.item.balanceColumn)+data.item.rowNumber; const inventoryRange=`${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!${balanceCell}`;
        await sheetsPut(`/${encodeURIComponent(state.workbookId)}/values/${encodeURIComponent(inventoryRange)}?valueInputOption=USER_ENTERED`,{range:inventoryRange,majorDimension:"ROWS",values:[[data.newBalance]]});
        closeConfirm(); $("movement-form").reset(); setDefaultTimestamp(); setMovementStatus("Movement recorded successfully."); pendingMovement=null; await loadLogger();
      }catch(e){console.error(e);setMovementStatus(e.message||"Unable to record movement.",true);}
      finally{$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
    }
  
    function closeConfirm(){$("confirm-modal").classList.add("hidden");pendingMovement=null;$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
    function readForm(){return{movement:$("movement").value,client:$("client").value.trim(),asset:$("asset").value,quantity:Number($("quantity").value),timestamp:$("timestamp").value};}
    function validateMovement(d){if(!d.movement||!d.client||!d.asset||!d.timestamp||!Number.isInteger(d.quantity)||d.quantity<=0)return"Please complete all fields and enter a whole quantity greater than zero.";return null;}
    function updatePreview(){const d=readForm();if(!d.movement||!d.client||!d.asset||!Number.isInteger(d.quantity)||d.quantity<=0){$("preview-text").textContent="Select the movement, client, asset and quantity.";return;}$("preview-text").textContent=`${movementLabel(d.movement)} ${formatNumber(d.quantity)} × ${d.asset} ${d.client?`for ${d.client}`:""}`;}
    function setDefaultTimestamp(){const now=new Date();$("timestamp").value=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;}
    function movementLabel(x){return x==="RECEIVED"?"Received":x==="SENT"?"Sent":"Discard";}
    function movementClass(x){return String(x||"").toLowerCase();}
    function formatTimestamp(x){if(!x)return"";const d=new Date(x);return Number.isNaN(d.getTime())?String(x):d.toLocaleString("en-GB",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"});}
    function formatNumber(x){return Number(x||0).toLocaleString("en-GB");}
    function numericValue(x){if(x===null||x===undefined||x==="")return 0;const n=Number(String(x).replace(/,/g,""));return Number.isFinite(n)?n:0;}
    function normalizeHeader(x){return String(x??"").trim().toLowerCase().replace(/\s+/g," ");}
    function findColumn(headers,names){for(const name of names){const idx=headers.indexOf(name);if(idx>=0)return idx;}return-1;}
    function quoteSheetName(x){return `'${String(x).replace(/'/g,"''")}'`;}
    function columnLetter(n){let r="";while(n>0){const rem=(n-1)%26;r=String.fromCharCode(65+rem)+r;n=Math.floor((n-1)/26);}return r;}
    function escapeDriveQuery(x){return String(x).replace(/\\/g,"\\\\").replace(/'/g,"\\'");}
    function escapeHtml(x){return String(x??"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}
    function escapeAttr(x){return escapeHtml(x);}
    function emptyRow(span,text){return `<tr><td colspan="${span}" class="empty">${escapeHtml(text)}</td></tr>`;}
    function authHeaders(){return{Authorization:`Bearer ${state.accessToken}`};}
    async function sheetsGet(path){return fetchJson(SHEETS_API+path,{headers:authHeaders()});}
    async function sheetsPost(path,body){return fetchJson(SHEETS_API+path,{method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}
    async function sheetsPut(path,body){return fetchJson(SHEETS_API+path,{method:"PUT",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}
    async function fetchJson(url,options){const response=await fetch(url,options);const text=await response.text();let data={};try{data=text?JSON.parse(text):{};}catch(_){}if(!response.ok)throw new Error(data?.error?.message||`Request failed (${response.status})`);return data;}
    function decodeJwtPayload(jwt){const parts=String(jwt).split(".");if(parts.length!==3)throw new Error("Invalid Google credential.");const base64=parts[1].replace(/-/g,"+").replace(/_/g,"/");const padded=base64+"=".repeat((4-base64.length%4)%4);return JSON.parse(decodeURIComponent(Array.from(atob(padded)).map(c=>`%${c.charCodeAt(0).toString(16).padStart(2,"0")}`).join("")));}
    function setAuthStatus(text,error=false){$("auth-status").textContent=text;$("auth-status").className=`status ${error?"error":""}`;}
    function setMovementStatus(text,error=false){$("movement-status").textContent=text;$("movement-status").className=`status ${error?"error":""}`;}
    function setSyncStatus(text,error=false){$("sync-status").textContent=text;$("sync-status").className=`muted ${error?"error":""}`;}
    function signOut(){if(state.idTokenPayload?.sub){try{google.accounts.id.revoke(state.idTokenPayload.sub,()=>{});}catch(_){} } state.idTokenPayload=null;state.accessToken=null;state.workbookId=null;state.workbookName=null;state.inventory=[];state.transactions=[];state.clients=[];$("logger").classList.add("hidden");$("login-card").classList.remove("hidden");$("google-signin-button").classList.remove("hidden");$("grant-access").classList.add("hidden");$("sign-out").classList.add("hidden");$("user-photo").classList.add("hidden");$("user-name").textContent="Not signed in";$("user-email").textContent="";}
  })();