const OfficeItemHandler = {
	get: function(target, prop, receiver) {
		// Console.log(`prop: ${prop.toString()}`);
		if (target === null ||
            target === undefined ||
            typeof target.Count !== 'number' ||
            typeof target.Item !== 'function')
            throw new TypeError('代理对象必须是一个带有Item方法的Office集合对象');
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