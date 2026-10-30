(()=>{const C=window.CMMS_CONFIG,$=s=>document.querySelector(s);
const show=()=>{const u=sessionStorage.cmmsUser;$("#login").hidden=!!u;$("#home").hidden=!u;if(u)$("#who").textContent=u};
$("#lf").onsubmit=async e=>{e.preventDefault();$("#go").disabled=true;$("#err").textContent="";
 try{const r=await fetch(C.url,{method:"POST",body:JSON.stringify({action:"login",username:$("#u").value.trim(),password:$("#p").value})});
  const j=await r.json();if(j.error)throw new Error(j.error);
  sessionStorage.cmmsToken=j.token;sessionStorage.cmmsUser=j.user;$("#p").value="";show()}
 catch(x){$("#err").textContent=/fetch|JSON/i.test(x.message)?"Cannot reach the server. Check the URL in config.js.":x.message}
 $("#go").disabled=false};
$("#out").onclick=()=>{sessionStorage.clear();show()};show()})();
