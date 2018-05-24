# ExcelFreedom
[ExcelFreedom.zip](https://github.com/nut077/ExcelFreedom/files/1987460/ExcelFreedom.zip)
<br>
[Tutorial.ExcelFreedom.pdf](https://github.com/nut077/ExcelFreedom/files/1987462/Tutorial.ExcelFreedom.pdf)

### Tutorial ExcelFreedom
รูปแบบการใช้งาน ```<table></table>``` คือ 1 sheet ```<tr></tr>``` คือ 1 แถว ```<td></td>``` คือ 1 column
<br>
การใช้งานแบบ offline
<br>
```java 
StringBuilder table = new StringBuilder();
table.append("<table>");
  table.append("<tr>");
    table.append("<td>A</td>");
    table.append("<td>B</td>");
    table.append("<td>C</td>");
  table.append("</tr>");
table.append("</table>");
// parameter ตัวแรกคือที่อยู่ของไฟล์ที่จะสร้าง ตัวที่สองคือชื่อของไฟล์ excel และตัวสุดท้ายคือข้อมูลที่จะเขียนลงในไฟล์ excel
ExcelFreedom excelFreedom = new ExcelFreedom("D://", "excel", table.toString());
excelFreedom.write(); // สั่งให้เขียนไฟล์
```
##### ผลลัพธ์ที่ได้คือ
[![Capture.png](https://s31.postimg.cc/oz427g9nf/Capture.png)](https://postimg.cc/image/5u0sxouzb/)
