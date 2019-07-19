module.exports = function(RED) {
    function excelTemplateFillingNode(config) {
        RED.nodes.createNode(this,config);
		this.dataForTemplate = config.dataForTemplate;
		this.templateExcelLocation = config.templateExcelLocation;
		this.reportExcelLocation = config.reportExcelLocation;
		this.sheetName = config.sheetName;
		this.columnDataTypes = config.columnDataTypes;
		this.headerRowNumber = config.headerRowNumber;
		this.writeType = config.writeType;
		
        var node = this;
        node.on('input', function(msg) {
        	var dataForTemplate = RED.util.evaluateNodetypeProperty(node.dataForTemplate,node,msg);
			var templateExcelLocation = RED.util.evaluateNodetypeProperty(node.templateExcelLocation,node,msg);
			var reportExcelLocation = RED.util.evaluateNodetypeProperty(node.reportExcelLocation,node,msg);
			var sheetName = RED.util.evaluateNodetypeProperty(node.sheetName,node,msg);
			var columnDataTypes = RED.util.evaluateNodetypeProperty(node.columnDataTypes,node,msg);
			var headerRowNumber = RED.util.evaluateNodetypeProperty(node.headerRowNumber,node,msg);
			var writeType = node.writeType;
			if(headerRowNumber != 'undefined'){
				headerRowNumber = parseInt(headerRowNumber);
			}else{
				headerRowNumber = 1;
			}
			
			if(typeof dataForTemplate == 'object'){
				dataForTemplate = JSON.stringify(dataForTemplate);
			}
			
			if(typeof columnDataTypes == 'object'){
				columnDataTypes = JSON.stringify(columnDataTypes);
			}
			
			if(reportExcelLocation == node.reportExcelLocation){
				reportExcelLocation = "";
			}
			var outputJSON = excelTemplateFilling(dataForTemplate, templateExcelLocation, reportExcelLocation, sheetName, columnDataTypes, headerRowNumber, writeType);
			
			console.log("OutPUT :::" + outputJSON);
			msg["result"] = JSON.parse(outputJSON);
			node.send(msg);
        });
    }
    RED.nodes.registerType("excel-template-filling",excelTemplateFillingNode);
	
}

function excelTemplateFilling(dataForTemplate, templateExcelLocation, reportExcelLocation, sheetName, columnDataTypes, headerRowNumber, writeType) {
	var java = require(process.env.NODE_RED_HOME+"/node_modules/java");
	var writeContenetInTemplate = java.import("com.hcl.ers.optimus.textprocessing.excel.WriteContentInTemplate");
	return writeContenetInTemplate.writeContentInTemplateSheetSync(dataForTemplate, templateExcelLocation, reportExcelLocation, sheetName, columnDataTypes, headerRowNumber, writeType);
}
