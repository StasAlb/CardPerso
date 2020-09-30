function show_user(par)
 {
    var ret = window.showModalDialog("UserAdd.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 440px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById("MyHiddenField").value = ret;
    if (ret != null) return true;
    else return false; 
 }
 function edit_user(par)
 {
    var ret = window.showModalDialog("UserEdit.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 440px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById("MyHiddenField").value = ret;
    if (ret != null) return true;
    else return false; 
 }
function edit_role(par)
 {
    var ret = window.showModalDialog("UserRole.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 440px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById("MyHiddenField").value = ret;
    if (ret != null)  return true;
    else return false; 
 } 
function show_role(par)
 {
    var ret = window.showModalDialog("RoleEdit.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 140px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById("MyHiddenField").value = ret;
    if (ret != null)  return true;
    else return false; 
 }
 
 function load_managerole() 
 {
    document.getElementById("pRoles").style.height=screen.height-345+'px';
 }
 function load_manageuser()
 {
    document.getElementById("pUsers").style.height = screen.height-345+'px';
 }
 function load_reminder() 
 {
    document.getElementById("pReminder").style.height=screen.height-340+'px';
 }
 function show_userfilter()
 {
    var ret = window.showModalDialog("FltUsers.aspx", 
    "dialogWidth: 500px; dialogHeight: 440px; center:yes; status:no; unadorned: yes; resizable: yes; help: no");
    document.getElementById("MyHiddenField").value = ret;
    if (ret != null) return true;
    else return false;
 }