 function show_branch(par)
 {
    var ret = window.showModalDialog("BranchEdit.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 310px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_stordoc(par)
 {
    var ret = window.showModalDialog("StorDocEdit.aspx?"+par, "", 
    "dialogWidth: 500px; dialogHeight: 310px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_product(par)
 {
    var ret = window.showModalDialog("ProductEdit.aspx?"+par, "", 
    "dialogWidth: 650px; dialogHeight: 240px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_addcomment(par)
 {
    var ret = window.showModalDialog("AddComment.aspx?"+par, "", 
    "dialogWidth: 400px; dialogHeight: 120px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_listdeliver(par)
 {
    var ret = window.showModalDialog("ListDeliverEdit.aspx?"+par, "", 
    "dialogWidth: 400px; dialogHeight: 180px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_productatt(par)
 {
    var ret = window.showModalDialog("ProductAttEdit.aspx?"+par, "", 
    "dialogWidth: 400px; dialogHeight: 180px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 function show_org(par)
 {
    var ret = window.showModalDialog("Catalog2Edit.aspx?"+par, "", 
    "dialogWidth: 520px; dialogHeight: 250px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 }
 function show_person(par)
 {
    var ret = window.showModalDialog("OrganizationEdit.aspx?"+par, "",
    "dialogWidth:520px; dialogHeight: 340px; center:yes; status:no; unadorned: yes; resizable: yes; help: no");
    document.getElementById('resd').value = ret;
    if (ret=="1") return true;
    else return false;
 }
 function show_operson(par)
 {
    var ret = window.showModalDialog("OPerson.aspx?"+par,"",
    "dialogWidth:520px; dialogHeight:340px, center:yes;status:no;unadored:yes; resizable:yes; help:no");
    document.getElementById('resde').value = ret;
    if (ret!=null) return true;
    else return false;
 }
 function show_field(par,dheight)
 {
    var ret = window.showModalDialog("ConfField.aspx?"+par, "", 
    "dialogWidth: 240px; dialogHeight:"+dheight+"; center:yes; status: no; unadorned: yes; resizable: no; help: no");         
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_purchase_product(par,type)
 {
    var ret = window.showModalDialog("PurchaseProductEdit.aspx?"+par, "", 
    "dialogWidth: 480px; dialogHeight:180px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    if (type=="1") return true;
    if (type=="2")
    {
        if (ret=="1") return true;
        else return false;
    }
 }
 
 function show_stordoc_product(par,type)
 {
    var ret = window.showModalDialog("StorDocProductEdit.aspx?"+par, "", 
    "dialogWidth: 450px; dialogHeight:180px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    if (type=="1") return true;
    if (type=="2")
    {
        if (ret=="1") return true;
        else return false;
    }
 }
 
 function show_stordoc_card(par,type)
 {
    var ret = window.showModalDialog("StorDocCardEdit.aspx?"+par, "", 
    "dialogWidth: 450px; dialogHeight:180px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    if (type=="1") return true;
    if (type=="2")
    {
        if (ret=="1") return true;
        else return false;
    }
 }
 
  function show_stordoc_cardrec(par)
 {
    var ret = window.showModalDialog("StorDocCardRec.aspx?"+par, "", 
    "dialogWidth: 700px; dialogHeight:560px; center:yes; status: no; unadorned: yes; resizable: no; help: no");         
    //"dialogWidth: 700px; dialogHeight:500px; center:yes; status: no; unadorned: yes; resizable: no; help: no");         
    if (ret=="1") return true;
    else return false; 
 }
 
 function show_purchase_dog(par)
 {
    var ret = window.showModalDialog("PurchaseDogEdit.aspx?"+par, "", 
    "dialogWidth: 800px; dialogHeight:250px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    if (ret=="1") return true;
    else return false;
 }
 
 function show_branch_office(par)
 {
    var ret = window.showModalDialog("BranchAddOffice.aspx?"+par, "", 
    "dialogWidth: 450px; dialogHeight:160px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    return true;
 } 
 function show_product_link(par)
 {
    var ret = window.showModalDialog("ProductLink.aspx?"+par, "", 
    "dialogWidth: 450px; dialogHeight:160px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    return true;
 } 

 function show_catalog(par)
 {
    var ret = window.showModalDialog("CatalogEdit.aspx?"+par, "", 
    "dialogWidth: 400px; dialogHeight: 120px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret=="1") return true;
    else return false; 
 } 
 function show_filedit()
 {
    var ret = window.showModalDialog("FilEdit.aspx?","", 
    "dialogWidth: 450px; dialogHeight: 120px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 
 function show_flt_storage()
 {
    var ret = window.showModalDialog("FltStorage.aspx?","", 
    "dialogWidth: 400px; dialogHeight: 200px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 
 function show_flt_purchase()
 {
    var ret = window.showModalDialog("FltPurchase.aspx?","", 
    "dialogWidth: 550px; dialogHeight: 240px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 
 function show_flt_stordoc()
 {
    var ret = window.showModalDialog("FltStorDoc.aspx?","", 
    "dialogWidth: 550px; dialogHeight: 260px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 
 function show_flt_card()
 {
    var ret = window.showModalDialog("FltCard.aspx?","", 
    "dialogWidth: 700px; dialogHeight: 730px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");
    HideWait();
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 function show_cardproperty(par)
 {
    var ret = window.showModalDialog("CardProperties.aspx?"+par, "", 
    "dialogWidth: 450px; dialogHeight:100px; center:yes; status: no; unadorned: yes; resizable: yes; help: no");         
    document.getElementById('resd').value = ret; 
    if (ret!=null && ret!="") return true;
    else return false; 
 }
 function show_date(par)
 {
    var ret = window.showModalDialog("ChangeDate.aspx?"+par, "",
    "dialogWidth:450px; dialogHeight:240px; center:yes; status:no; unadored:yes; resizable: yes; help:no");
    document.getElementById('resdd').value = ret;
    if (ret!=null && ret!="") return true;
    else return false;    
 }
 function load_organization()
 {
    document.getElementById("pOrg").style.height = screen.height-340+'px';
 }

 function load_branch() 
 {
    //alert(request.querystring('mode'));
  //  if (document.getElementById("pBranch222")!=null)
  //      document.getElementById("pBranch222").style.height=screen.height-295+'px';
    document.getElementById("pBranch").style.height=screen.height-340+'px';
 }
 
 function load_supplier() 
 {
    document.getElementById("pSupplier").style.height=screen.height-340+'px';
 }
 
 function load_manufacturer() 
 {
    document.getElementById("pManufacturer").style.height=screen.height-340+'px';
 }

 function load_courier() 
 {
    document.getElementById("pCourier").style.height=screen.height-340+'px';
 }
 
 function load_bank() 
 {
    document.getElementById("pBank").style.height=screen.height-340+'px';
 }
 
function load_expendables()
{
    document.getElementById("pExpendable").style.height=screen.height-340+'px';
}
 
 function load_product() 
 {
    document.getElementById("pProduct").style.height=screen.height-340+'px';
 }
 
 function load_listdeliver() 
 {
    document.getElementById("pListDeliver").style.height=screen.height-340+'px';
 }
 
 function load_productatt() 
 {
    document.getElementById("pProductAtt").style.height=screen.height-340+'px';
 }
 
 function load_storage() 
 {
    document.getElementById("pStorage").style.height=screen.height-340+'px';
 }
 
 function load_purchase() 
 {
    document.getElementById("pDogs").style.height=(screen.height-340)/2+'px';
    document.getElementById("pProducts").style.height=(screen.height-340)/2+'px';
 }
 
 function load_stordoc() 
 {
    document.getElementById("MySplitter").style.height = screen.height-340+'px';
//    document.getElementById("pDocs").style.height=(screen.height-320)/2+'px';
//    document.getElementById("pProducts").style.height=(screen.height-320)/2+'px';
//    document.getElementById("pCards").style.height=(screen.height-320)/2+'px';
 } 
 
function SetToolTip(image, hintName) {
    var ret = window.showModalDialog("ShowCard.aspx?im="+image, "",
    "dialogWidth:505px; dialogHeight:665px; center:yes; status:yes; unadored:yes; resizable: yes; help:yes");
    return true;

//     var str = "<img width='480px' height='640px' src='" + image + "'/>";
//     //var str = "<img src='" + image + "'/>";
//     tooltip(this, str, hintName)
}

