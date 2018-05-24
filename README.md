# ExcelFreedom
[ExcelFreedom.zip](https://github.com/nut077/ExcelFreedom/files/1987460/ExcelFreedom.zip)
<br>
[Tutorial.ExcelFreedom.pdf](https://github.com/nut077/ExcelFreedom/files/1987462/Tutorial.ExcelFreedom.pdf)

### Tutorial ExcelFreedom
* รูปแบบการใช้งาน<br>
```<table></table>``` คือ 1 sheet ```<tr></tr>``` คือ 1 แถว ```<td></td>``` คือ 1 column
```java StringBuilder table = new StringBuilder();
table.append("<table>");
  table.append("<tr>");
    table.append("<td>A</td>");
    table.append("<td>B</td>");
    table.append("<td>C</td>");
  table.append("</tr>");
table.append("</table>");

// ใช้งานแบบ offline
// parameter ตัวแรกคือที่อยู่ของไฟล์ที่จะสร้าง ตัวที่สองคือชื่อของไฟล์ excel และตัวสุดท้ายคือข้อมูลที่จะเขียนลงในไฟล์ excel
ExcelFreedom excelFreedom = new ExcelFreedom("D://", "excel", table.toString());

// ใช้งานแบบ online ไฟล์ jsp
// parameter ตัวแรกคือ response ตัวที่สองคือ out ตัวที่สามคือชื่อของไฟล์ excel และตัวสุดท้ายคือข้อมูลที่จะเขียนลงในไฟล์ excel
ExcelFreedom excelFreedom = new ExcelFreedom(response, out, "excel", table.toString());

excelFreedom.write(); // สั่งให้เขียนไฟล์
```
##### ผลลัพธ์ที่ได้คือ
[![Capture.png](https://s31.postimg.cc/oz427g9nf/Capture.png)](https://postimg.cc/image/5u0sxouzb/)
<br>

* เปลี่ยนชื่อ sheet
โดยค่าเริ่มต้นจะตั้งชื่อเป็น sheet 1

