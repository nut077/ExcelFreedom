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
    <br><br>

* เพิ่ม sheet<br>
    เพิ่ม tag ขึ้นมาใหม่ก็จะได้ sheet ใหม่ขึ้นมา จำนวน sheet ขึ้นอยู่กับ tag <br>
    ตัวอย่างในที่นี้เราจะสร้าง 2 sheet
    ```java StringBuilder table = new StringBuilder();
        table.append("<table>");
          table.append("<tr>");
            table.append("<td>A</td>");
            table.append("<td>B</td>");
            table.append("<td>C</td>");
          table.append("</tr>");
        table.append("</table>");
        table.append("<table>");
          table.append("<tr>");
            table.append("<td>D</td>");
            table.append("<td>E</td>");
            table.append("<td>F</td>");
          table.append("</tr>");
        table.append("</table>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/cp66feunj/Capture.png)](https://postimg.cc/image/70zvoiqaz/)
    <br>
