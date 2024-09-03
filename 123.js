function CommandButton1_Click() {
	//初始化处理，设定读取列范围和列号
	let wb = ThisWorkbook;
	let sht1 = wb.Sheets(1);
	const fangwei = [3, 6];
	const columns = [4, 5, 6];
	//con

	//调用函数读取D列到F列的数据
	let SheetsValue = SheetValue(sht1, fangwei);
	if (!SheetsValue) return MsgBox("表格数据不能为空", 0 + 48, "警告");

	//调用函数做合并单元格预处理
	let CellsValue = checkcells(sht1, columns, SheetsValue);

	//调用自定义计算函数，计算G列和F列的值
	let GFValue = jsq(sht1, CellsValue);

	//输出G列和F列的数据
	Output(sht1, GFValue);
	return alert("计算完成");
	//	Console.log(JSON.stringify(GFValue));//调试用
}


//获取表格数据，提取其中指定列的数据（可以处理空行数据）
function SheetValue(sht, arry) {
	let results = [];
	//使用 UsedRange 获取工作表中所有已使用的范围，不论数据是否连续
	let sheetData = sht.UsedRange.Value2;
	//排除第一个元素（子数组）包含列头
	for (let i = 1; i < sheetData.length; i++) {
		//提取每行的D到F列数据
		let rowData = sheetData[i].slice(arry[0], arry[1]);
		//使用正则表达式来检查提取的行是否只包含空白字符
		if (!/^[\s]*$/.test(rowData.join(''))) {
			results.push(rowData);
		}
	}
	return results.length > 0 ? results : "";
}


//检查指定列当前行是否为合并单元格，是则循环找指定列上一行中不为空的变量的值，否则直接使用
function checkcells(sheet, arry, data) {
	let results = [];
	//一层循环遍历每行数据，二层循环遍历这行的指定列
	for (let i = 0; i < data.length; i++) {
		let CellsValue = [];
		for (let k = 0; k < arry.length; k++) {
			//从第2行开始检查，排除表头
			let cell = data[i][k];
			if (sheet.Cells(i + 2, arry[k]).MergeCells && isNaN(data[i][k])) {
				let found = false; // 声明一个标记变量，用来标记是否找到非NaN的值（即非空值）
				for (let j = i - 1; j >= 0 && !found; j--) {
					if (!isNaN(data[j][k])) { // 如果找到一个非NaN的值
						cell = data[j][k]; // 更新cell的值
						found = true; // 更新标记变量，退出循环
					}
				}
			}
			CellsValue.push(Number(cell));
			Console.log(JSON.stringify(CellsValue));
		}
		results.push(CellsValue);
		//  		Console.log(JSON.stringify(results));
	}
	return results;
}


//浮点计算G列和F列的值，得出的结果返回
function jsq(sheet, data) {
	let results = [];
	let rows = data.length;
	for (let i = 0; i < rows; i++) {
		//赋值并做预处理，转换为浮点数
		let dValue = parseFloat(data[i][0]);
		let eValue = parseFloat(data[i][1]);
		let fValue = parseFloat(data[i][2]);
		//计算结果
		let gValue = dValue * fValue;
		let hValue = (eValue * gValue) / 1000;
		//分别检查G列和H列计算结果是否有值，否则提示错误
		if (isNaN(gValue)) {
			gValue = "数据缺失";
		}
		if (isNaN(hValue)) {
			hValue = "数据缺失";
		}
		//使用push方法将2个元素做为一个元素（子数组）添加进数组results里
		results.push([gValue, hValue]);
	}
	return results;
}


//遍历数组每一行，将各行第一个数据输出到G列，第2个输出到H列
function Output(sheet, data) {
	for (let i = 0; i < data.length; i++) {
		//设定输出数据的起始行
		let j = i + 2;
		sheet.Cells(j, 7).Value2 = data[i][0];
		sheet.Cells(j, 8).Value2 = data[i][1];
	}
}