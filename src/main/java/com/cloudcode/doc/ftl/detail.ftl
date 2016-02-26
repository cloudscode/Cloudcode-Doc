<!DOCTYPE html>
<html lang="en">
<head>
   <#include "classpath:com/cloudcode/framework/common/ftl/head.ftl"/>
    <style type="text/css">
select {font-family:Verdana,sans-serif,宋体;width:150px;}	
  </style>
   
</head>
<body data-spy="scroll" data-target=".bs-docs-sidebar" data-twttr-rendered="true"> 
<div id="dialogDiv">
<div class="container" id="layout">
<form role="form" class="form-horizontal" id="myFormId" action="${request.getContextPath()}/futurestypes/createFuturesType" method="post">
   <div class="form-group">
    <label for="inputEmail3" class="col-sm-2 control-label">交易所</label>
    <div class="col-sm-4">     
      <select name="exchange" id="exchange">
      <option  selected="selected" value="1">大连商品交易所</option>
      <option value="2">郑州商品交易所</option>
      <option value="3">上海期货交易所</option>
      <option value="4">中国金融期货交易所</option>
      <option value="5">渤海商品交易所</option>
      <option value="6">泛亚有色金属交易所</option>
    </select>
    </div>
     <label for="inputPassword3" class="col-sm-2 control-label">类别</label>
    <div class="col-sm-4">
      <select name="groupType" id="groupType">
      <option  selected="selected" value='1'>PVC</option>
      <option value='2'>棕油</option>
      <option value='3'>豆二</option>
      <option value='4'>豆粕</option>
      <option value='5'>铁矿石</option>
      <option value='6'>鸡蛋</option>
      <option value='7'>塑料</option>
      <option value='8'>PP</option>
      <option value='9'>纤维板</option>
      <option value='10'>胶合板</option>
      <option value='11'>豆油</option>
      <option value='12'>玉米</option>
      <option value='13'>豆一</option>
      <option value='14'>焦炭</option>
      <option value='15'>焦煤</option>
      <option value='16'>玉米淀粉</option>
    </select>
	</div>
    </div>
   <div class="form-group">
    <label for="inputEmail3" class="col-sm-2 control-label">编号</label>
    <div class="col-sm-10">
      <input type="text" name="code" class="form-control" id="code" placeholder="编号">
    </div>
  </div>  
  <div class="form-group">
    <div class="col-sm-offset-2 col-sm-10">
     <input type="hidden" value="" id="oid" name="id" >
       <button type="button" id="updateButton" class="btn btn-default">save</button>
    </div>
  </div>
</form>

</div>
<#include "classpath:com/cloudcode/framework/common/ftl/vendor.ftl"/>
<script type="text/javascript">

var hm = $("body").wHumanMsg();
$(function () {
$("#exchange").change(function(e,o){
var r = $(this).children('option:selected').val();

$("#groupType").empty();
if(r=="1"){
  $("#groupType").append("<option  selected='selected' value='1'>PVC</option>"); 
       $("#groupType").append("<option value='2'>棕油</option>"); 
       $("#groupType").append("<option value='3'>豆二</option>"); 
       $("#groupType").append("<option value='4'>豆粕</option>"); 
       $("#groupType").append("<option value='5'>铁矿石</option>"); 
       $("#groupType").append("<option value='6'>鸡蛋</option>"); 
        $("#groupType").append("<option value='7'>塑料</option>"); 
         $("#groupType").append("<option value='8'>PP</option>"); 
          $("#groupType").append("<option value='9'>纤维板</option>"); 
           $("#groupType").append("<option value='10'>胶合板</option>"); 
            $("#groupType").append("<option value='11'>豆油</option>"); 
             $("#groupType").append("<option value='12'>玉米</option>"); 
              $("#groupType").append("<option value='13'>豆一</option>"); 
               $("#groupType").append("<option value='14'>焦炭</option>"); 
                $("#groupType").append("<option value='15'>焦煤</option>"); 
                 $("#groupType").append("<option value='16'>玉米淀粉</option>"); 
}else
	if(r=="2"){         
		$("#groupType").append("<option value='17'>PTA</option>"); 
		$("#groupType").append(" <option value='18'>菜油</option>"); 
		$("#groupType").append("<option value='19'>菜籽</option>"); 
		$("#groupType").append(" <option value='20'>菜粕</option>"); 
		$("#groupType").append(" <option value='21'>动力煤</option>"); 
		$("#groupType").append("<option value='22'>强麦</option>"); 
		$("#groupType").append("<option value='23'>粳稻</option>"); 
		$("#groupType").append("  <option value='24'>白糖</option>"); 
		$("#groupType").append("<option value='25'>棉花</option>"); 
		$("#groupType").append("<option value='26'>早籼稻</option>"); 
		$("#groupType").append(" <option value='27'>甲醇</option>"); 
		$("#groupType").append("<option value='28'>晚籼稻</option>"); 
		$("#groupType").append("<option value='29'>硅铁</option>"); 
		$("#groupType").append("<option value='30'>锰硅</option>"); 
	
	}else if(r=="3"){
	$("#groupType").append("<option value='31'>燃油</option>"); 
    $("#groupType").append("<option value='32'>沪铝</option>"); 
    $("#groupType").append("<option value='33'>橡胶</option>"); 
    $("#groupType").append("<option value='34'>沪锌</option>"); 
    $("#groupType").append("<option value='35'>沪铜</option>"); 
    $("#groupType").append("<option value='36'>黄金</option>"); 
                           $("#groupType").append("<option value='37'>螺纹钢</option>"); 
                            $("#groupType").append("<option value='38'>线材</option>"); 
                            $("#groupType").append(" <option value='39'>沪铅</option>"); 
                              $("#groupType").append("<option value='40'>白银</option>"); 
                     $("#groupType").append("<option value='41'>沥青</option>"); 
                      $("#groupType").append("<option value='42'>热轧卷板</option>"); 
                       $("#groupType").append("<option value='43'>沪锡</option>"); 
                        $("#groupType").append("<option value='44'>沪镍</option>"); 
	}else if(r=="4"){
	 $("#groupType").append("<option value='45'>期指</option>"); 
     $("#groupType").append(" <option value='46'>5年期国债期货</option>"); 
     $("#groupType").append("<option value='47'>10年期国债期货</option>");                         
                            
	}else if(r=="5"){
	   $("#groupType").append("<option value='48'>现货商品</option>"); 
	}else if(r=="6"){
	    $("#groupType").append("<option value='49'>有色金属</option>"); 
	}
});
    //####### Buttons
   // $("#layout button,.button,#sampleButton").button();
   if('${entityAction}' =='update'){
	   $('#updateButton').click( function() {
				     if($('#selectTextVal').val() ==''){
				     	$('#selectTextVal').val('root');
				     }
	      		   $.ajax({
				        url: '${request.getContextPath()}/futurestypes/'+$("#oid").val()+'/updateFuturesType',
				        type: 'post',
				        dataType: 'json',
				        data:$('form#myFormId').serialize(),
				        success: function(data) {
				       		 $('body').wHumanMsg('theme', 'black').wHumanMsg('msg', '数据保存成功！', {fadeIn: 300, fadeOut: 300});
				        	 $('.ui-dialog-titlebar-close').trigger('click');
				           }
			    });
		});
   }else{
	    $('#updateButton').click( function() {
			  
			    $.ajax({
			        url: '${request.getContextPath()}/futurestypes/createFuturesType',
			        type: 'post',
			        dataType: 'json',
			        data: $('form#myFormId').serialize(),
			        success: function(data) {
			       		 $('body').wHumanMsg('theme', 'black').wHumanMsg('msg', '数据保存成功！', {fadeIn: 300, fadeOut: 300});
			       		 $('.ui-dialog-titlebar-close').trigger('click');
			         }
			    });
		});
    }
	$("#modal-message").dialog({
	    autoOpen: false,
	    modal: true,
	    buttons: {
	        Ok: function () {
	        	//$('#divInDialog').dialog("close");
	            $(this).dialog("close");
	           
	        }
	    }
	});
	if('${entityAction}' =='update'){
		<#if menu?exists>  
		$('#exchange').val('${futuresType.exchange!''}');
		$('#groupType').val('${futuresType.groupType!''}');
		$('#code').val('${futuresType.code!''}');
		$('#oid').val('${futuresType.id!''}');
		</#if>
	}
});
</script>
</div>
</body>
</html>