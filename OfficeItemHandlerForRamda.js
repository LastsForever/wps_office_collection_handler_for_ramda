//  通过对带有Item方法的WPS中的Office集合对象添加Handler，使得Ramda库的R.map与R.filter能够直接在Office对象上操作。
//  使用前需：
//  	(1) 新建一个WPS的JS宏代码模块，将Ramda库的源码复制到其中(dist/ramda.js文件);
//      (2) 再与同一文件内新建一个代码模块，将本文件代码全部复制到其中;
//	(2) WPS JS宏编辑器中，工具 => 选项 => 编译 => 取消“禁用全局作用域表达式”及“禁用全局作用域标识符重复定义”;
//  其中，Ramda库版本： v0.28.0
//  链接： https://github.com/ramda/ramda
//  

const OfficeItemHandler = {
	get: function(target, prop, receiver) {
		// Console.log(`prop: ${prop.toString()}`);
		if (target === null ||
            target === undefined ||
            typeof target.Count !== 'number' ||
            typeof target.Item !== 'function')
            throw new TypeError('必须是一个带有Item方法的Office集合对象');
        else if (prop === 'length')
         	return target.Count;
        else if (prop === 'map')
        	return function(fn) {
        		let arr = [];
        		for (let item of target)
        			arr.push(fn(item));
        		return arr;
        	}
        else if (prop.constructor.name === 'String' && /^\d+$/.test(prop))
		return target.Item(Number(prop) + 1);
	else
		return Reflect.get(target, prop, receiver);
	}
}

// 测试：对Range对象调用ramda的map和filter功能，找出加粗的单元格并输出其地址。
function ramda_test() {
	var rng = new Proxy(ActiveSheet.UsedRange, OfficeItemHandler);
	R.pipe(
		R.filter(cell => cell.Font.Bold === true),
		R.map(cell => cell.Address()),
		arr => Console.log(arr.join(";")),
	)(rng);
}
