//  通过对 WPS Office集合对象（需带有Item方法）使用代理，使之可被Ramda库的R.map与R.filter直接操作。
//  使用方法：
//  	(1) 新建一个WPS JS宏代码模块，复制Ramda源码(dist/ramda.js文件)至其中;
//      (2) 于同一文件内新建代码模块，将本文件代码复制其中;
//	(3) WPS JS宏编辑器中，工具 => 选项 => 编译 => 分别取消“禁用全局作用域表达式”和“禁用全局作用域标识符重复定义”;
//	(4) 为WPS Office集合对象新建代理（见ramda_test函数）;
//  注：
//  	Ramda版本： v0.28.0
//  	链接： https://github.com/ramda/ramda
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
	let office_collection_object = ActiveSheet.UsedRange;
	let proxy_range_object = new Proxy(office_collection_object, OfficeItemHandler);
	R.pipe(
		R.filter(cell => cell.Font.Bold === true),
		R.map(cell => cell.Address()),
		arr => Console.log(arr.join(";")),
	)(proxy_range_object);
}
