# ExcelFreedom
[ExcelFreedom.zip](https://github.com/nut077/ExcelFreedom/files/1987460/ExcelFreedom.zip)
<br>
[Tutorial.ExcelFreedom.pdf](https://github.com/nut077/ExcelFreedom/files/1987462/Tutorial.ExcelFreedom.pdf)

### Tutorial ExcelFreedom
* **รูปแบบการใช้งาน**<br>
```<table></table>``` คือ 1 sheet ```<tr></tr>``` คือ 1 แถว ```<td></td>``` คือ 1 column
    ```java
    StringBuilder table = new StringBuilder();
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

* **เพิ่ม sheet**<br>
เพิ่ม tag ```<table></table>``` ขึ้นมาใหม่ก็จะได้ sheet ใหม่ขึ้นมา จำนวน sheet ขึ้นอยู่กับ tag ```<table>```<br>
ตัวอย่างในที่นี้เราจะสร้าง 2 sheet
    ```java
    StringBuilder table = new StringBuilder();
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
    <br><br>
    
* **เปลี่ยนชื่อ sheet**<br>
ใช้ tag ```<sheet>ชื่อ sheet</sheet>``` โดยใส่ไว้ต่อจาก ```<table>```
    ```java
    table.append("<table>");
    table.append("<sheet>new sheet</sheet>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/m8kgmxzxb/Capture.png)](https://postimg.cc/image/jr8pfog0r/)
    <br><br>
    
* **ผสานเซลล์ แนวนอน**<br>
ใช้ tag ```<colspan>ตัวเลขที่ต้องการ</colspan>```
    ```java
    table.append("<td><colspan>2</colspan>colspan 2</td>");
    table.append("<td><colspan>3</colspan>colspan 3</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/6ea5qgdbj/Capture.png)](https://postimg.cc/image/49pspdbor/)
    <br><br>
    
* **ผสานเซลล์ แนวตั้ง**<br>
ใช้ tag ```<rowspan>ตัวเลขที่ต้องการ</rowspan>```
    ```java
    table.append("<td><rowspan>2</rowspan>rowspan 2</td>");
    table.append("<td><rowspan>3</rowspan>rowspan 3</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/tk8m5p2ez/Capture.png)](https://postimg.cc/image/bu6xknotz/)
    <br><br>
    
* **ขยายขนาดความกว้างของคอลัมน์**<br>
ใช้ tag ```<width>ตัวเลขที่ต้องการ</width>```
    ```java
    table.append("<td><width>10</width>width 10</td>");
    table.append("<td><width>20</width>width 20</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/5j1q4e0i3/Capture.png)](https://postimg.cc/image/nym71semf/)
    <br><br>
    
* **ขยายขนาดความสูงของคอลัมน์**<br>
ใช้ tag ```<height>ตัวเลขที่ต้องการ</height>```
    ```java
    table.append("<tr>");
        table.append("<td><height>500</height>height 500</td>");
    table.append("</tr>");
    table.append("<tr>");
        table.append("<td><height>600</height>height 600</td>");
    table.append("</tr>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/6uvdd4tfj/Capture.png)](https://postimg.cc/image/4qb0c1rsr/)
    <br><br>
    
* **การจัดรูปแบบ**<br>
ใช้ tag ```<format>รูปแบบตาม list ข้างล่าง</format>``` ค่าเริ่มต้นคือ border-center
    - left
    - center
    - right
    - left-middle
    - center-middle
    - right-middle
    - left-top
    - center-top
    - right-top
    - orientation
    - orientation-middle
    - orientation-top
    ```java
    table.append("<td><format>left</format>left</td>");
    table.append("<td><format>center</format>center</td>");
    table.append("<td><format>right</format>right</td>");
    table.append("<td><format>left-middle</format>left-middle</td>");
    table.append("<td><format>center-middle</format>center-middle</td>");
    table.append("<td><format>right-middle</format>right-middle</td>");
    table.append("<td><format>left-top</format>left-top</td>");
    table.append("<td><format>center-top</format>center-top</td>");
    table.append("<td><format>right-top</format>right-top</td>");
    table.append("<td><format>orientation</format>orientation</td>");
    table.append("<td><format>orientation-middle</format>orientation-middle</td>");
    table.append("<td><format>orientation-top</format>orientation-top</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/f8yrbc2zv/Capture.png)](https://postimg.cc/image/85qvvpxk7/)
    <br><br>
