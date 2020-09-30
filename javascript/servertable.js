(function($,window,document) 
				{
									
									function fnServerTable() 
									{
										this._this=this;
										
										this._PrepareRequestData=function(oSet)
										{
											var req=[];
											req.push({name:"sTableId",value:oSet.sId});
											req.push({name:"nPageRow",value:oSet.nPageRow});
											req.push({name:"nPageCurrent",value:oSet.nPageCurrent});
											req.push({name:"nSortColumnIndex",value:oSet.aSort[0]-((oSet.fnIsSelect()==true)?1:0)});
											req.push({name:"sSortColumnDir",value:oSet.aSort[1]});
											req.push({name:"isSelect",value:oSet.fnIsSelect()});
											return req;
										};
										//-----------------------------------------------------------------------------
										this._SelectRow=function(oSet,col,value)
										{
											var bFind=false;
											var that=this;
											if(oSet.fnIsSelect()==false) return true;
											if(typeof(col)=='string' && col=='FIRST_ROW')
											{
												var chk=$('input',$("td",$("tr",$(oSet.nTBody)).first()).first());
												if(chk.length)
												{	
													//if(!chk.attr("checked")) $(chk).trigger("click");
													if(!chk.attr("checked")) $(chk).attr("checked",true);
													that._IsSelectingRows(oSet);
													bFind=true;
												}	
											}	
											else
											{
												var oTr=$("tr",$(oSet.nTBody));
												oTr.each(function(iRow)
												{
													var oTd=$("td",this);
													var oTrThis=this;
													oTd.each(function(iCol)
													{
														if(iCol-1==col)
														{
															var txt=$('div',this).text();
															if(txt=='' + value) 
															{
																var chk=$('input',$("td",oTrThis)).first();
																if(chk.length)
																{	
																	//if(!chk.attr("checked")) $(chk).trigger("click");
																	if(!chk.attr("checked")) $(chk).attr("checked",true);
																	that._IsSelectingRows(oSet);
																	bFind=true;
																}
															}	
														}	
													});
												});	
												
											}	
											return bFind;
										};
										//-----------------------------------------------------------------------------
										
										this._ClearData=function(oSet,aSort)
										{
											oSet.nPageCurrent=0;
											if(typeof(aSort)=='undefined') oSet.aSort=[-1,'none'];
											this._FillRows(oSet,[]);	
										};
										//-----------------------------------------------------------------------------
										this._RefreshToFirstDesc=function(oSet)
										{
											oSet.nPageCurrent=0;
											if(oSet.aoColumns[(oSet.fnIsSelect())?1:0].bSortable==true) oSet.aSort=[(oSet.fnIsSelect())?1:0,'desc'];
											this._SortingClasses(oSet);
											this._RequestData(oSet);
										};
										//-----------------------------------------------------------------------------
										this._RequestData=function(oSet)
										{
											jsonReq=this._PrepareRequestData(oSet);
											if(oSet.fnPrerequest(jsonReq)==false) return;
											oSet.blockClick=true;
											oSet.fnShowWait();
											var that=this;
											$.ajax({
												"async":true,
												"type":"POST",
												"url":oSet.url,
												"data":jsonReq,
												"success": function ok(json)	
												 {	
												    oSet.fnHideWait();

												    if (json != null && typeof (json) != 'undefined' && typeof (json.SESSION_END) != 'undefined' && json.SESSION_END == true)
												    {
												        oSet.blockClick = false;
												        if (typeof (SESSION_END) == 'function') SESSION_END();
												        return;
												    }

													var bOK=that._GetRows(oSet,json); 
													oSet.blockClick=false; 
													setTimeout(function() { oSet.fnEndDataLoad(bOK); },100); 
												 },
												"dataType":"json",
												"cache": false,
												"error": function (xhr, error, thrown) 
												 {
													oSet.fnHideWait();
													var err="Данные для таблицы " + oSet.sId + " не получены.";
													if(error=="parsererror" ) 
													{
														oSet.fnShowError(err + " Ошибка при разборе JSON." );
													}
													else oSet.fnShowError(err + " Ошибка при обращении к " + oSet.url);
													oSet.blockClick=false;
													setTimeout(function() { oSet.fnEndDataLoad(false); },100);
												 }
											});
										};
										//-----------------------------------------------------------------------------
										this._GetRows=function(oSet,jsonData) 
										{
											//var oData = (typeof $.parseJSON == 'function') ? $.parseJSON(jsonData.replace(/'/g, '"') ) : eval('('+jsonData+')');
											if(jsonData.sTableId!=oSet.sId) 
											{ 
												oSet.fnShowError("Полученный идентификатор таблицы \"" + jsonData.sTableId + "\" не совпадает с \"" + oSet.sId + "\"."); 
												return false; 
											}
											if(typeof(jsonData.error)!='undefined') 
											{
												oSet.fnShowError(jsonData.error); 
												return false; 	
											}	
												
											if(typeof(jsonData.aRows)!='undefined' && jsonData.aRows.length>0)
											{
												for(var i=0;i<jsonData.aRows.length;i++)
												{	
											
													if(jsonData.aRows[i].length!=oSet.aoColumns.length)
													{
														oSet.fnShowError("Несовпадение количества колонок таблицы \"" + jsonData.sTableId + "\" в полученных строках.");
														return false;
													}
												}
											}
											this._FillRows(oSet,jsonData);
											return true;
										};
										//-----------------------------------------------------------------------------
										this._FillRows=function (oSet,jsonData)
										{
											var that=this;
											var dataLen=(typeof(jsonData.aRows)!='undefined' && jsonData.aRows.length>0) ? jsonData.aRows.length:0;
											var oTBody=oSet.nTBody;
											if(oTBody==null) return;
											if(oSet.selectable=='all')
											{	
												var oTh=$("th",$(oSet.nTHead));
												$('input',$(oTh)).first().attr("checked",false);
											}
											var oTr=$("tr",$(oTBody));
											oSet.nRowsShowPage=dataLen;
											oTr.each(function(iRow)
											{
												var oTd=$("td",this);
												oTd.each(function(iCol)
												{
													var curCol=iCol-((oSet.fnIsSelect()) ? 1:0);
													var bAttr=false;
													if(curCol>=0 && oSet.bColumnsTitle.length>0 && curCol<oSet.bColumnsTitle.length) bAttr=oSet.bColumnsTitle[curCol];
													var bHtml=false;
													if(curCol>=0 && oSet.bColumnsHtml.length>0 && curCol<oSet.bColumnsHtml.length) bHtml=oSet.bColumnsHtml[curCol];
													if(iRow<dataLen)
													{
														$(this).css('border-right-width','1px').css('border-bottom-width','1px');
														$(this).css('background',oSet.sBackGround).css('color',oSet.sColor);
														if(iCol==0 && oSet.fnIsSelect())
														{	
															$(this).empty();
															var chkbox=$(oSet.sCheckbox);
															$(this).append(chkbox);
															if(oSet.fnIsSelectRow())
															{	
																$(chkbox).css('display','none');
																$(this).css('border-right-width','0');
															}	
															$(this).append('<span style="display:none;">' + jsonData.aRows[iRow][iCol] + '</span>');
															$(chkbox).bind('change',function c() 
															{ 
																if(oSet.blockClick) return;
																if(oSet.selectable=='one' || oSet.selectable=='row')
																{
																	if($(this).attr("checked"))
																	{	
																		var id=$(this).next().text();
																		$("tr",$(oSet.nTBody)).each(function(iRow)
																		{
																	  		var oTd=$("td",this).first();
																	    	var chk=$('input',$(oTd)).first();
																	    	if($(chk).next().text()!=id)
																	    	{	
																				if(chk.length) $(chk).attr("checked",false);
																	    	}	
																		});
																	}	
																}
																that._IsSelectingRows(oSet);
															});
															
														}
														else
														{	
															if(bAttr==false || jsonData.aRows[iRow][iCol].length<1) $(this).removeAttr("title");
															else $(this).attr("title",jsonData.aRows[iRow][iCol]);
															if(bHtml) $('div',this).html(jsonData.aRows[iRow][iCol]);
															else $('div',this).text(jsonData.aRows[iRow][iCol]);
														}	
														$(this).css('width',parseInt(oSet.aoColumns[iCol].sWidth)+'px');
													}
													else
													{	
														if(iCol<oSet.aoColumns.length-1) $(this).css('border-right-width','0px');
														if(iRow!=oSet.nPageRow-1) $(this).css('border-bottom-width','0px');
														if(iCol==0 && oSet.fnIsSelect())
														{	
															$(this).empty();
														}
														else $('div',this).text('');
														$(this).removeAttr("title");
													}
												});
											});
											var nRows=0;
											if(typeof(jsonData.nRows)!='undefined') nRows=jsonData.nRows;
											if(typeof(jsonData.nPageCurrent)!='undefined') oSet.nPageCurrent=jsonData.nPageCurrent;
											oSet.nPageCount=Math.floor(nRows/oSet.nPageRow) + ((nRows%oSet.nPageRow==0) ? 0:1);
											oSet.nRows=nRows;
											oSet.nRowsAll=0;
											if(typeof(jsonData.nRowsAll)!='undefined') oSet.nRowsAll=jsonData.nRowsAll;
											this._IsSelectingRows(oSet);
											this._FillPages(oSet);
										};
										//-----------------------------------------------------------------------------
										this._FillPages=function (oSet)
										{
											if(oSet.nList==null) return;
											var that=this;
											var nList=oSet.nList;
											$(nList).empty();
											if(oSet.nPageCount!=0)
											{	
												var pageMax=(oSet.nPageDisplay>oSet.nPageCount) ? oSet.nPageCount:oSet.nPageDisplay;
												var shiftCentr=(Math.floor(pageMax/2) + ((pageMax%2==0) ? 0:1));
												var shiftPage=oSet.nPageCurrent-shiftCentr;
												if(shiftPage<0) shiftPage=0;
												else
												{	
													if(shiftPage+pageMax>=oSet.nPageCount)	shiftPage=oSet.nPageCount-pageMax;
												}	
												for(var i=shiftPage;i<pageMax+shiftPage;i++)
												{
													var span=null;
													if(oSet.nPageCurrent!=i)
													{
														span=$('<span class="'+ oSet.sPageButton + '">'+ (i+1) + '</span>');
														$(nList).append(span).css('vertical-align','middle');
														span.css("border-width",'0px');
														span.click(function c() {
															if(oSet.blockClick) return;
															var p=($(this).text());
															oSet.nPageCurrent=parseInt(p)-1;
															that._RequestData(oSet);
														});
													}
													else
													{
														span=$('<span class="'+ oSet.sPageButtonDisable + '">' + (i+1) + '</span>');
														$(nList).append(span).css('vertical-align','middle');
														span.css("border-width",'1px');
													}
													span.width(span.width()+4).height(span.height()+2);
												}
											}
											else
											{
												var span=$('<span class="'+ oSet.sPageButtonDisable + '">-</span>');
												$(nList).append(span).css('vertical-align','middle');
												span.css("border-width",'0px');
											}	
											var rClass=oSet.sPageButton + " " + oSet.sPageButtonDisable;
											$(oSet.nFirst).removeClass(rClass);
											$(oSet.nPrevious).removeClass(rClass);
											$(oSet.nNext).removeClass(rClass);
											$(oSet.nLast).removeClass(rClass);
											if(oSet.nPageCount==0)
											{	
												$(oSet.nFirst).addClass(oSet.sPageButtonDisable);
												$(oSet.nPrevious).addClass(oSet.sPageButtonDisable);
												$(oSet.nNext).addClass(oSet.sPageButtonDisable);
												$(oSet.nLast).addClass(oSet.sPageButtonDisable);
											}
											else
											{	
											  if(oSet.nPageCurrent==0) 
											  {
													$(oSet.nFirst).addClass(oSet.sPageButtonDisable);
													$(oSet.nPrevious).addClass(oSet.sPageButtonDisable);
											  }
											  else
											  {
													$(oSet.nFirst).addClass(oSet.sPageSeek);
													$(oSet.nPrevious).addClass(oSet.sPageSeek);
											  }
											  if(oSet.nPageCurrent==oSet.nPageCount-1)
											  {
												  $(oSet.nNext).addClass(oSet.sPageButtonDisable);
												  $(oSet.nLast).addClass(oSet.sPageButtonDisable);
											  }
											  else
											  {
												  $(oSet.nNext).addClass(oSet.sPageSeek);
												  $(oSet.nLast).addClass(oSet.sPageSeek);
											  }	  
												  
											}
											$(oSet.nRec).text(this._GetTextRecords(oSet));
											$(oSet.nRec).attr('title',this._GetTextRecords(oSet));
										};
										//-----------------------------------------------------------------------------
										this._IsSelectingRows=function(oSet)
										{
											if(oSet.fnIsSelect()==false) return;
											var that=this;
											var checkrows=[];
											var cRow=0;
											$("tr",$(oSet.nTBody)).each(function(iRow)
											{
												 var oTd=$("td",this).first();
												 var chk=$('input',$(oTd));
												 if(oSet.fnIsSelectRow()==true)
												 {	  
													 $(this).css('cursor','default'); 
													 $("td",this).each(function(iCol) { if(iCol!=0) $(this).unbind('click');});
												 }	  
												 if(chk.length)
												 {	 
													if(oSet.fnIsSelectRow()==true)
													{	 
													 $("td",this).each(function(iCol) 
													 { 
														if(iCol!=0)
														{	 
														 $(this).click('click',function() 
														 { 
															 if(!chk.attr("checked")) $(chk).attr("checked",true);
															 else $(chk).attr("checked",false);
															 $(chk).trigger("change");
														 });
														} 
													 });
													}
													 
													 cRow++;
													 if($(chk).attr("checked"))
													 {
														 $("td",this).each(function(iCol) 
														 {
															 $(this).css('background',oSet.sBackGroundSelect);
															 $('div',this).css('background',oSet.sBackGroundSelect).css('font-weight','bold');
															 if(oSet.sBackGroundSelect!=oSet.sBackGround && iCol<oSet.aoColumns.length-1) $(this).css('border-right-color',oSet.sBackGroundSelect);
														 });
														 checkrows.push($('span',$(oTd)).text());							 
													 }
													 else
													 {	 
														 $("td",this).each(function(iCol) 
														 {
															 $(this).css('background',oSet.sBackGround);
															 $('div',this).css('background',oSet.sBackGround).css('font-weight','normal');
															 if(oSet.sBackGroundSelect!=oSet.sBackGround && iCol<oSet.aoColumns.length-1) $(this).css('border-right-color',oSet.border_right_color);
															 });
														 }	 
													 }
												 else
												 {
													 $("td",this).each(function(iCol) 
													 {
														 $(this).css('background',oSet.sBackGround);
														 $('div',this).css('background',oSet.sBackGround).css('font-weight','normal');
														 if(oSet.sBackGroundSelect!=oSet.sBackGround && iCol<oSet.aoColumns.length-1) $(this).css('border-right-color',oSet.border_right_color);
													 });	
												 }	 
											
											});
											if(checkrows.length<1) 
											{
												if(oSet.selectable=='all') $('input',$(oSet.nThead).first()).first().attr("checked",false);			
											}
											else
											{
												if(oSet.selectable=='all' && checkrows.length==cRow) $('input',$(oSet.nThead).first()).first().attr("checked",true);
											}
											//сюда пользовательскую функцию для передачи выделенных записей
											setTimeout(function() {	oSet.fnSelectRows(checkrows); }, 100);
										};
										//-----------------------------------------------------------------------------
										this._SelectAll=function(oSet,check)
										{
											if(oSet.selectable!='all') return;
											var oTh=$("tr",$(oSet.nTHead));
											$('input',$(oTh)).first().attr("checked",check);
											$("tr",$(oSet.nTBody)).each(function(iRow)
											{
												 var oTd=$("td",this).first();
												 var chk=$('input',$(oTd));
												 if(chk.length) $(chk).attr("checked",check);
											});
											this._IsSelectingRows(oSet);
										};
										//-----------------------------------------------------------------------------
										this._AddColumnSelect=function(oSet,nTh)
										{
											var that=this;
											var aoColumns=[];
											aoColumns.push({
													"bSortable": false,
													"sSortingClass": oSet.sSortClassNone,
													"sTitle": '',
													"nTh": document.createElement('th'),
													"sWidth":'16px'
											});
											for(var i=0;i<oSet.aoColumns.length;i++) aoColumns.push(oSet.aoColumns[i]);  
											oSet.aoColumns=aoColumns;
											var iCol=0;
											var oCol = oSet.aoColumns[iCol];
											var nth=oSet.aoColumns[iCol].nTh;
											nTh.parentNode.insertBefore(nth,nTh);
											if(oSet.selectable=='all')
											{	
												var chkbox=$(oSet.sCheckbox);
												$(nth).append(chkbox);
												$(chkbox).bind('change',function c() 
												{ 
													if(oSet.blockClick) return;
													if($(this).attr('checked')) that._SelectAll(oSet,true);
													else that._SelectAll(oSet,false);
												}
												);
											}
											else
											{
												if(oSet.fnIsSelectRow()==false) $(nth).text('W');
												$(nth).width(oSet.aoColumns[iCol].sWidth);
												$(nth).css('color',$(nth).css('background-color'));
											}
											$(nth).css('vertical-align','middle');
											$(nth).attr('align','center');
											
											//if(oSet.sBackGroundSelect==null) oSet.sBackGroundSelect=$(nth).css('background');
											
											if(oSet.sBackGroundSelect==null) oSet.sBackGroundSelect=oSet.sBackGround;
											
											oCol.sWidth=$(nth).width();
											if(oCol.sWidth<1) oCol.sWidth=16; 
											if(oSet.fnIsSelectRow())
											{	
												oCol.sWidth=0; 
												$(nth).width(0);
												$(nth).css('border-right-width','0');
												
											}	
										};
										//-----------------------------------------------------------------------------
										this._AddColumn=function(oSet,nTh)
										{
											var that=this;
											oSet.aoColumns[oSet.aoColumns.length++] = {
												"bSortable": true,
												"sSortingClass": oSet.sSortClassNone,
												"sTitle": nTh ? nTh.innerHTML : '',
												"nTh": nTh ? nTh : document.createElement('th'),
												"sWidth":'0px'		
											};
											var iCol=oSet.aoColumns.length-1;
											var oCol = oSet.aoColumns[iCol];
											if(oSet.bColumnsSort.length>0 && iCol<oSet.bColumnsSort.length) oCol.bSortable=oSet.bColumnsSort[iCol];
											/*
											var nDiv = document.createElement('div');
											nDiv.className=""; //oSet.sSortClassNone;
											var nth=oSet.aoColumns[iCol].nTh;
											$(nth).contents().appendTo(nDiv);
											nDiv.appendChild(document.createElement('span'));
											nth.appendChild(nDiv);
											$('span',$(nth)).css('float','right');
											*/
											
											var nDiv=document.createElement('div');
											var nT=document.createElement('table');
											var nR=document.createElement('tr');
											nR.appendChild(document.createElement('td'));
											nR.appendChild(document.createElement('td'));
											nT.appendChild(nR);
											nDiv.appendChild(nT);
											$(nT).attr('border','0').attr('cellspacing','0').attr('cellspadding','0').attr('width','100%');
											$(nR).attr('valign','middle');
											nDiv.className="";
											var nth=oSet.aoColumns[iCol].nTh;
											$(nth).contents().appendTo($('td',$(nDiv)).first());
											($('td',$(nDiv)).last()).append(document.createElement('span'));
											$('td',$(nDiv)).first().attr('align','center').css('white-space','nowrap');
											$('td',$(nDiv)).last().attr('align','right');
											nth.appendChild(nDiv);
											$(nDiv).css('overflow','hidden');
											var w=parseInt($(nth).attr('width'));
											if(w>$(nDiv).width()) $(nDiv).width(w); 
											$(nth).click(function c() 
											{ 
												var index=-100;
												if(oSet.blockClick) return;
												for(var i=0;i<oSet.aoColumns.length;i++) if(oSet.aoColumns[i].nTh==this) { index=i; break; }
												if(index>=0) that._ClickToSort(oSet,index);	
											});
											$(nth).css('vertical-align','middle');//.css('white-space','nowrap');
											oCol.sWidth=$(nth).width();
										};
										//-----------------------------------------------------------------------------
										this._ReplaceCSSHeader=function(oSet)
										{
											var oTHead=oSet.nTHead;
											if(oTHead==null) return;
											var oTh=$("th",$(oTHead));
											
											oTh.each(function(iCol)
											{
												if(iCol>0) $(this).css('border-left-width','0px');
												$(this).css('padding','2px 2px 2px 2px').css('vertical-align','middle');
											});
										};
										//-----------------------------------------------------------------------------
										this._ClickToSort=function(oSet,index)
										{
											if(oSet.aoColumns[index].bSortable==false) return;
											if(oSet.aSort[0]==index)
											{
												if(oSet.aSort[1]!="asc") oSet.aSort[1]="asc";
												else oSet.aSort[1]="desc";
											}
											else 
											{
												oSet.aSort[0]=index;
												oSet.aSort[1]="asc";
											}
											this._SortingClasses(oSet);
											this._RequestData(oSet);
										};
										//-----------------------------------------------------------------------------
										this._SortingClasses=function(oSet)
										{
											for(var i=0;i<oSet.aoColumns.length;i++)
											{
												$(oSet.aoColumns[i].nTh).removeClass(oSet.sSortClassNone);
												$(oSet.aoColumns[i].nTh).css('cursor','default');
												if(oSet.aoColumns[i].bSortable)
												{	
													var jqSpan = $("span",oSet.aoColumns[i].nTh);
													jqSpan.removeClass(oSet.sSortClass +" "+ oSet.sSortClassDesc +" " + oSet.sSortClassAsc);
													var sSpanClass=oSet.sSortClass;
													if(oSet.aSort[0]==i)
													{	
														if(oSet.aSort[1]=="asc")	sSpanClass=oSet.sSortClassAsc;
														else
														if(oSet.aSort[1]=="desc")	sSpanClass=oSet.sSortClassDesc;	
													}
													jqSpan.addClass(sSpanClass);
													jqSpan.css('float','right');
													
													
												}
												$(oSet.aoColumns[i].nTh).addClass(oSet.sSortClassNone);

											}
										};	
										//-----------------------------------------------------------------------------
										this._CreateRows=function(oSet)
										{
											if(oSet.nTBody!=null)
											{	
												$(oSet.nTBody).empty();
												for(var r=0;r<oSet.nPageRow;r++)
												{	
													var nTr=document.createElement('tr');
													for(var i=0;i<oSet.aoColumns.length;i++)
													{
														var nTd=document.createElement('td');
														
																												
														//nTd.innerText=" ";
														nTr.appendChild(nTd);
														
														$(nTd).addClass(oSet.sClassBase + ' ' + oSet.sClassDefault);
														$(nTd).css('background',oSet.sBackGround).css('color',oSet.sColor).css('font-weight','normal')
															.css('padding','2px 2px 2px 2px').css('white-space','nowrap').css('overflow','hidden');
																												
														if(i<oSet.aoColumns.length-1)
														{	
															 $(nTd).css('border-right-width','0').css('border-top-width','0');
															 if(i!=0) $(nTd).css('border-left-width','0');
														}	
														else
														{	
															$(nTd).css('border-left-width','0').css('border-top-width','0');
														}
														if(r!=oSet.nPageRow-1) { $(nTd).css('border-bottom-width','0'); }
														if(oSet.fnIsSelect() && i==0) $(nTd).attr('align','center');
																																					
														$(nTd).attr('width',parseInt(oSet.aoColumns[i].sWidth)+'px');
														
														var nD=document.createElement('div');
														nTd.appendChild(nD);
														
														$(nD).text(" ");
														//var nth=oSet.aoColumns[i].nTh;
														$(nD).width(oSet.aoColumns[i].sWidth);
														$(nD).css('background',oSet.sBackGround).css('color',oSet.sColor).css('font-weight','normal')
															.css('white-space','nowrap').css('overflow','hidden');
														
														
													}
													oSet.nTBody.appendChild(nTr);
													$(nTr).addClass(oSet.sClassBase);
													var h=$(nTr).css('font-size');
													h=parseInt(h);
													h=$(nTr).height(h+8);
												}
											}
										};
										//-----------------------------------------------------------------------------
										this._GetTextRecords=function(oSet)
										{
											var txt=(oSet.nRows<1)?"Показано " + oSet.nRows:"Показано: " + (oSet.nPageCurrent*oSet.nPageRow+1) +
																				" - " + ((oSet.nPageCurrent*oSet.nPageRow+1)+oSet.nRowsShowPage-1) +
																				" из " + oSet.nRows;
											if(oSet.nRows!=oSet.nRowsAll && oSet.nRowsAll>0) txt+=" ( всего " + oSet.nRowsAll +" )";
											txt+="      ";
											return txt;
										};
										//-----------------------------------------------------------------------------
										this._SelectPageButton=function(oSet,oButton)
										{
											var page=-1;
											if(oSet.nPageCount>0)
											{
												if(oButton==oSet.nFirst)
												{
													if(oSet.nPageCurrent>0) page=0;
												}
												if(oButton==oSet.nPrevious)
												{
													if(oSet.nPageCurrent>0) page=oSet.nPageCurrent-1;
												}
												if(oButton==oSet.nNext)
												{
													if(oSet.nPageCurrent<oSet.nPageCount-1) page=oSet.nPageCurrent+1;
												}
												if(oButton==oSet.nLast)
												{
													if(oSet.nPageCurrent<oSet.nPageCount-1) page=oSet.nPageCount-1;
												}
											}
											if(page>=0)
											{
												oSet.nPageCurrent=page;
												this._RequestData(oSet);		
											}	
										};
										//-----------------------------------------------------------------------------
										this._CreatePaging=function(oSet)
										{
											var that=this;
											var nPageDiv=document.createElement('div');
											var nRec=document.createElement('span');
											var nFirst=document.createElement('span');
											var nPrevious=document.createElement('span');
											var nList=document.createElement('span');
											var nNext=document.createElement('span');
											var nLast=document.createElement('span');

											nFirst.title='Начало';
											nPrevious.title="Назад";
											nNext.title="Дальше";
											nLast.title="Конец";
											
															
											nPageDiv.className=oSet.sPageButton;
											nRec.className=oSet.sPageButton;
											nFirst.className=oSet.sPageButtonDisable + " " + oSet.sPageFirst;
											nPrevious.className=oSet.sPageButtonDisable + " " + oSet.sPagePrevious;
											nNext.className=oSet.sPageButtonDisable + " " + oSet.sPageNext;
											nLast.className=oSet.sPageButtonDisable + " " + oSet.sPageLast;

											nPageDiv.appendChild(nRec);
											nPageDiv.appendChild(nFirst);
											nPageDiv.appendChild(nPrevious);
											nPageDiv.appendChild(nList);
											nPageDiv.appendChild(nNext);
											nPageDiv.appendChild(nLast);
											

											$('span',nPageDiv).bind( 'mousedown', function () { return false; } ).bind( 'selectstart', function () { return false; });
											$('span',nPageDiv).css('border-width','0px');
											
											if(oSet.idInsertAfter==null)
											{	
												var parent=oSet.nTable.parentNode;
												if(parent.lastchild==oSet.nTable.parentNode) 
												{
													parent.appendChild(nPageDiv);
												} 
												else 
												{
													parent.insertBefore(nPageDiv,oSet.nTable.nextSibling);
												}
											}
											else 
											{
												$(nPageDiv).insertAfter('#' + oSet.idInsertAfter);
												//$(nPageDiv).css('float','left')
												//$(nPageDiv).css('display','inline-block');
											}	
										    $('span',nPageDiv).css('vertical-align','middle');
										    
										    $(nPageDiv).css('margin','4px 4px 0px 0px');
										    $(nPageDiv).css('padding','4px 4px 4px 4px');
										    $(nRec).css('cursor','default');
										    
										    $(nFirst).click(function c() 	{ if(oSet.blockClick!=true) { that._SelectPageButton(oSet,this);} });
										    $(nPrevious).click(function c() { if(oSet.blockClick!=true) { that._SelectPageButton(oSet,this);} });
										    $(nNext).click(function c() 	{ if(oSet.blockClick!=true) { that._SelectPageButton(oSet,this);} });
										    $(nLast).click(function c() 	{ if(oSet.blockClick!=true) { that._SelectPageButton(oSet,this);} });
										    
										    oSet.nPageDiv=nPageDiv;
											oSet.nRec=nRec;
										    oSet.nFirst=nFirst;
											oSet.nPrevious=nPrevious;
											oSet.nList=nList;
											oSet.nNext=nNext;
											oSet.nLast=nLast;
											
											$(nRec).text(this._GetTextRecords(oSet));
											$(nRec).attr('title',this._GetTextRecords(oSet));
											
										};

									};
									
									var methods=
									{
						   			 init:function(args) 
						    		 {
										args=args || {};
										//-----------------------------------------------------------------------------
										function oSetClass()
										{
											this.selectable=args.selectable || "none"; // one,all,none,row
											if(this.selectable!='all' && this.selectable!='one' && this.selectable!='row') this.selectable='none';
											this.fnShowError=args.fnShowError || function(msg) { alert(msg); };
											this.url=args.url || "url_not_defined";
											this.aSort=args.aSort || [-1,'none'];
											this.fnShowWait=args.fnShowWait || function() {  };
											this.fnHideWait=args.fnHideWait || function() {  };
											this.nPageRow=args.nPageRow || 10; 
											this.nPageDisplay=args.nPageDisplay || 14;
											this.sBackGround=args.sBackGround || 'white';
											this.sBackGroundSelect=args.sBackGroundSelect || null;
											this.sColor=args.sColor || 'black';
											this.fnPrerequest=args.fnPrerequest || function(json) { return true; };
											this.fnSelectRows=args.fnSelectRows || function(ids)  { };
											this.idInsertAfter=args.idInsertAfter || null;
											this.bColumnsSort=args.bColumnsSort || []; // Флаги сортировки колонок (по умолчанию все сортируемые)
											this.bFirstLoad=(typeof(args.bFirstLoad)!='undefined') ? args.bFirstLoad:true;
											this.fnEndDataLoad=args.fnEndDataLoad || function(success)  { };
											this.bColumnsTitle=args.bColumnsTitle || []; // Флаги для вывода титлов для ячеек (по умолчанию их нет)
											this.bColumnsHtml=args.bColumnsHtml || []; // Флаги для вывода ячеек в колонках, как html (по умолчанию text)
											//-------------------------------------------------------------------------
											this.nPageCount=0;
											this.nPageCurrent=0;
											this.nRows=0;
											this.nRowsAll=0;
											this.nRowsShowPage=0;
											//-------------------------------------------------------------------------
											this.sId="";
											this.bInit=false;
											this.nTable=null;
											this.aoColumns=[];
											this.nThead=null;
											this.sSortClassNone='ui-state-default';
											this.sSortClass='ui-icon ui-icon-carat-2-n-s';
											this.sSortClassAsc='ui-icon ui-icon-carat-1-n';
											this.sSortClassDesc='ui-icon ui-icon-carat-1-s';
											this.sClassBase='ui-widget';
											this.sClassDefault='ui-state-default';
											this.nTHead=null;
											this.nTBody=null;
											//-------------------------------------------------------------------------
											this.sPageButton="ui-button ui-state-default";
											this.sPageButtonDisable="ui-button ui-state-disabled";
											this.sPageFirst="ui-icon ui-icon-seek-start";
											this.sPagePrevious="ui-icon ui-icon-seek-prev";
											this.sPageNext="ui-icon ui-icon-seek-next";
											this.sPageLast="ui-icon ui-icon-seek-end";
											this.sPageSeek="ui-button ";
											this.nPageDiv=null;
											this.nRec==null;
											this.nFirst==null;
											this.nPrevious==null;
											this.nList==null;
											this.nNext==null;
											this.nLast==null;
											this.fnIsSelect=function() { return (this.selectable!='none'); };
											this.sCheckbox='<input type="checkbox" class="ui-widget-content ui-corner-all"/>';
											this.blockСlick=false;
											this.fnIsSelectRow=function() { return (this.selectable=='row'); };
											this.border_right_color=null;
											if(this.aSort[0]>=0) this.aSort[0]+=(this.fnIsSelect()==true)?1:0;
										}
										//-----------------------------------------------------------------------------
										return this.each(function()
										{
													var data = $(this).data('serverTable');
													if(data!=null)
													{
														var oSet=data.oSet;
														if(oSet.bInit==true) return;
													}	
													var curSet=new oSetClass();
													var sT=new fnServerTable();
													var sId=this.getAttribute('id');
													if(sId==null ) { curSet.fnShowError("Не указан Id таблицы"); return; }
													curSet.sId=sId;
													if(this.nodeName.toLowerCase()!='table') { curSet.fnShowError("Не найден <table id=" + sId + ">"); return; }
													curSet.nTable=this;
													var nThead=this.getElementsByTagName('thead');
													if(nThead==null || nThead.length!=1)  { curSet.fnShowError("Элемент <thead> в <table id=" + sId + "> должен быть один"); return; }
													curSet.nThead=nThead[0];
													var nTrs = nThead[0].getElementsByTagName('tr');
													if(nTrs==null || nTrs.length!=1)  { curSet.fnShowError("Элемент <tr > в <thead> в <table id=" + sId + "> должен быть один"); return; }
													var anThs=nTrs[0].getElementsByTagName('th');
													if(anThs!=null && anThs.length>0)	
													{ 
														for(var i=0;i<anThs.length; i++ )
														{	
															sT._AddColumn(curSet,anThs[i]);
														}
														if(curSet.fnIsSelect()) sT._AddColumnSelect(curSet,anThs[0]);
													}
													else sT._AddColumn(curSet,null);
													$(nTrs).addClass(curSet.sClassBase);
													$(nTrs).height($(nTrs).height() + 16);
													
													//$('div',$(nTrs)).height($(nTrs).height());
													
													if(this.getElementsByTagName('tbody').length === 0 ) this.appendChild(document.createElement('tbody'));
													curSet.nTHead = this.getElementsByTagName('thead')[0];
													curSet.nTBody = this.getElementsByTagName('tbody')[0];
													sT._ReplaceCSSHeader(curSet);
													sT._SortingClasses(curSet);
													sT._CreateRows(curSet);
													sT._CreatePaging(curSet);
													
													if(curSet.border_right_color==null)
													{	
														var oTh=$("th",$(curSet.nTHead));
														curSet.border_right_color=$(oTh).css("border-right-color");
													}	
													
													var wTable=0;
													for(var i=0;i<curSet.aoColumns.length;i++)
													{
														wTable+=parseInt(curSet.aoColumns[i].sWidth);
													}
																										
													//$('#' + sId).css('table-layout','fixed');
													$('#' + sId).attr('width',wTable+'px');
																										
										            
										            $(this).data('serverTable',{oSet:curSet,fn:sT});
										          
										            if(curSet.bFirstLoad==true) sT._RequestData(curSet);
										            
													curSet.bInit=true;
											});
									},
									refresh:function()
									{
										return this.each(function()
										{
											var data = $(this).data('serverTable');
								            if(data)
								            {
								            	var oSet=data.oSet;
								            	data.fn._RequestData(oSet);
								            }
										});
									},
									clear:function()
									{
										return this.each(function()
										{
											var data = $(this).data('serverTable');
								            if(data)
								            {
								            	var oSet=data.oSet;
								            	data.fn._ClearData(oSet);
								            }
										});
									},
									tofirstdesc:function()
									{
										return this.each(function()
										{
											var data = $(this).data('serverTable');
								            if(data)
								            {
								            	var oSet=data.oSet;
								            	data.fn._RefreshToFirstDesc(oSet);
								            }
										});
									},
									selectrow:function(col,value)
									{
										var data = $(this).data('serverTable');
										if(typeof(col)=='undefined')
						            	{
						            		col='FIRST_ROW';value='';
						            	}
										if(data)
							            {
							            		
							            	var oSet=data.oSet;
							            	return data.fn._SelectRow(oSet,col,value);
							            }
							            return false;
									},
									getsort:function()
									{
										var data = $(this).data('serverTable');
								        if(data)
								        {
								           	var oSet=data.oSet;
								           	return [oSet.aSort[0]-((oSet.fnIsSelect()==true)?1:0),oSet.aSort[1]];
								        }
								        else return [-1,'none']; 
									},
									isdatapresent: function ()
									{
									    var data = $(this).data('serverTable');
									    if (data)
									    {
									        var oSet = data.oSet;
									        return (oSet.nRows>0);
									    }
									    else return false;
									}
								};
								$.fn.serverTable = function(method)
								{
									
								    if(methods[method] )
								    {
								      	return methods[method].apply(this,Array.prototype.slice.call(arguments,1));
								    } 
								    else 
								    if(typeof(method)==='object' || !method) 
								    {
								      	return methods.init.apply(this,arguments);
								    } 
								    else 
								    {
								      $.error('Метод ' +  method + ' в jQuery.serverTable не существует');
								    }    
								   		
								};
				})(jQuery, window, document);