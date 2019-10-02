# k-excel

##### 导出传参：需要导出的数据，数据长度，设置（每份表格的条数，是否压缩，表头）.
##### 通过让客户端选择是否分表导出，自动整合数据导出.
##### 导入传参：绑定的组件'element'如ice的上传组件，导入的字段，回调函数.

> 导出参考
```
import KExcel from 'k-excel';
.
.
.
cosnt data =[{employer_id: 11, employer_name: "11"},{employer_id: 12, employer_name: "12"}];
const page = {total: 12, end: 2}
const toExcel = new KExcel();
toExcel.excelZip(data, page,
  {
	sheetLength: 3,
	isZip: -1,
	sheet: [
	  { key: 'employer_id', value: 'id' },
	  { key: 'employer_name', value: '名称' },
	],
  },
);
```
 * 此外，若数据是对象里含有对象，即{A{B}}的形式，key需要以A.B的形式写入
 
### 导入参考
```
import KExcel from 'k-excel';
.
.
.
const pullExcel = new KExcel();
pullExcel.excelPull(element, [
  { key: 'salary_name' },
  { key: 'salary_id' },
],, then);

```
> UPDATE
 * 1.1.2 修复解析字符串的问题
 * 1.1.0 增加了自定义导入方法
 * 1.0.9 support ie
 * 1.0.8 按客户端设置的分表数量导出
 * 1.0.7 修复数据为奇数时，缺条数的bug
 * 1.0.6 较为完善版本
 * send advice to 450947795@qq.com \ ^_^ /