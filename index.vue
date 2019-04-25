<template>
	<div>
		<div class="controlBar clearfix">
			<el-button class='right' type="success" icon="el-icon-circle-plus-outline" @click="goExcel">导入文件</el-button>
         		<el-button @click="importExcel">确定导入</el-button>
        	</div>
	</div>
</template>
<script>
import $ from 'jquery'
import XLSX from 'xlsx'
export default {
	data(){
		return {
			
		}
	},
	methods:{
		importAjax(url,data){
            		this.$axios({
                		method: "post",
                		url: url,
                		contentType: 'json',
                		data:data,
            		}).then(response => {
            			if(response.data.msg == "success"){
                			this.$message.success("导入成功!");
            			}else if(response.data.msg == "error"){
                			this.$message.error("导入失败");
            			}else{
                			this.$message.error("导入失败，"+response.data.msg);
            			}
            		}).catch(err => {
            			console.log(err);
            		});
        	},
        	//excel文件判断
        	isExcel() {
            		var fileDir = $("#excel-file").val();
            		var suffix = fileDir.substr(fileDir.lastIndexOf("."));
           		if (fileDir =="") {
                		return this.$message.error("请先选择需要导入的excel文件!");
            		}
            		if(suffix != ".xls" && suffix != ".xlsx") {
                		return this.$message.error("文件类型不正确,请选择excel文件!");
            		}
            		return true;
        	},
        	//ecxel文件解析
        	importExcel() {
            		if(this.isExcel()){
                		var files = $("#excel-file")[0].files;
                		var fileReader = new FileReader();
                		fileReader.onload = (ev)=>{
                    			try {
                        			var data = ev.target.result,
                        			workbook = XLSX.read(data, {
                            			type: 'binary' // 以二进制流方式读取得到整份excel表格对象
                        		}),
                        		persons = []; // 存储获取到的数据
                    			} catch (e) {
                        			return this.$message.error("文件类型不正确,请选择excel文件!");
                    			}
                    			// 表格的表格范围，可用于判断表头是否数量是否正确
                    			var fromTo = '';
                    			// 遍历每张表读取
                    			for (var sheet in workbook.Sheets) {
                        			if (workbook.Sheets.hasOwnProperty(sheet)) {
                            				fromTo = workbook.Sheets[sheet]['!ref'];
                            				var sheet = sheet;//表头
                            				persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                            				// break; // 如果只取第一张表，就取消注释这行
                        			}
                    			}
                    			var data = [];
                    			for (var i=0; i<persons.length; i++){
                        			if (i >= 2) data.push(persons[i]);
                    			}
                		};
                		this.importAjax('/XXXXX/test/importExcel',{'NAME':persons});
                		// 以二进制方式打开文件
                		fileReader.readAsBinaryString(files[0]);
            		}
        	},	
	}
}
</script>
