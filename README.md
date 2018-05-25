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
    [![Capture.png](https://s31.postimg.cc/fqtwr70aj/Capture.png)](https://postimg.cc/image/lrrlo9mwn/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์ทั้งหมด ใช้ border
    - border-left
    - border-center
    - border-right
    - border-left-middle
    - border-center-middle
    - border-right-middle
    - border-left-top
    - border-center-top
    - border-right-top
    - border-orientation
    - border-orientation-middle
    - border-orientation-top

    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/5qtc18gnz/Capture.png)](https://postimg.cc/image/r0gyc2wyj/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์บนและล่าง ใช้ border-tb
    - border-tb-left
    - border-tb-center
    - border-tb-right
    - border-tb-left-middle
    - border-tb-center-middle
    - border-tb-right-middle
    - border-tb-left-top
    - border-tb-center-top
    - border-tb-right-top
    - border-tb-orientation
    - border-tb-orientation-middle
    - border-tb-orientation-top
    
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/sjcchgo7z/Capture.png)](https://postimg.cc/image/iyspukyvv/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์บนและล่างแบบ 2 เส้น ใช้ border-tb-double
    - border-tb-double-left
    - border-tb-double-center
    - border-tb-double-right
    - border-tb-double-left-middle
    - border-tb-double-center-middle
    - border-tb-double-right-middle
    - border-tb-double-left-top
    - border-tb-double-center-top
    - border-tb-double-right-top
    - border-tb-double-orientation
    - border-tb-double-orientation-middle
    - border-tb-double-orientation-top

    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/ir52xvdkv/Capture.png)](https://postimg.cc/image/p4u614igr/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์บน 1 เส้น และ ล่างแบบ 2 เส้น ใช้ border-tb-single-double
    - border-tb-single-double-left
    - border-tb-single-double-center
    - border-tb-single-double-right
    - border-tb-single-double-left-middle
    - border-tb-single-double-center-middle
    - border-tb-single-double-right-middle
    - border-tb-single-double-left-top
    - border-tb-single-double-center-top
    - border-tb-single-double-right-top
    - border-tb-single-double-orientation
    - border-tb-single-double-orientation-middle
    - border-tb-single-double-orientation-top

    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/5d29phzh7/Capture.png)](https://postimg.cc/image/jji0kqac7/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์เฉพาะล่างแบบ 2 เส้น ใช้ border-b-double
    - border-b-double-left
    - border-b-double-center
    - border-b-double-right
    - border-b-double-left-middle
    - border-b-double-center-middle
    - border-b-double-right-middle
    - border-b-double-left-top
    - border-b-double-center-top
    - border-b-double-right-top
    - border-b-double-orientation
    - border-b-double-orientation-middle
    - border-b-double-orientation-top

    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/t5bl14we3/Capture.png)](https://postimg.cc/image/cubh4tjw7/)
    <br><br>
    ใส่กรอบให้กับคอลัมน์ทั้งหมดเฉพาะล่างแบบ 2 เส้น ใช้ border-all-b-double
    - border-all-b-double-left
    - border-all-b-double-center
    - border-all-b-double-right
    - border-all-b-double-left-middle
    - border-all-b-double-center-middle
    - border-all-b-double-right-middle
    - border-all-b-double-left-top
    - border-all-b-double-center-top
    - border-all-b-double-right-top
    - border-all-b-double-orientation
    - border-all-b-double-orientation-middle
    - border-all-b-double-orientation-top

    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/5fm5cbvez/Capture.png)](https://postimg.cc/image/u8vpczefb/)
    <br><br>
    ถ้าต้องการให้กรอบเป็นเส้นปะให้ใช้ dashed โดยเอามาวางต่อจาก border<br>
    ตัวอย่างเช่น border-dashed-center
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/idhj3f5bj/Capture.png)](https://postimg.cc/image/x9g2b0gq3/)
    <br><br>
    
* **เปลี่ยน font**<br>
ใช้ tag ```<font-name>รูปแบบตาม list ข้างล่าง<font-name>``` ค่าเริ่มต้นคือ arial
    - arial
    - tahoma
    - courier
    - times
    - thsarabun_spk
    - thsarabun_new
    ```java
    table.append("<td><font-name>arial</font-name>arial</td>");
    table.append("<td><font-name>tahoma</font-name>tahoma</td>");
    table.append("<td><font-name>courier</font-name>courier</td>");
    table.append("<td><font-name>times</font-name>times</td>");
    table.append("<td><font-name>thsarabun_spk</font-name>thsarabun_spk</td>");
    table.append("<td><font-name>thsarabun_new</font-name>thsarabun_new</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/sbmgj8uy7/Capture.png)](https://postimg.cc/image/fwzoix3fv/)
    <br><br>
    
* **เปลี่ยนขนาด font**<br>
ใช้ tag ```<font-size>ขนาด<font-size>``` ค่าเริ่มต้นคือ 10
    ```java
    table.append("<td><font-size>10</font-size>font default</td>");
    table.append("<td><font-size>20</font-size>font size 20</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/cdii0oqq3/Capture.png)](https://postimg.cc/image/4kru8pkqv/)
    <br><br>
    
* **ตัวอักษรตัวหนา**<br>
ใช้ tag ```<b>true or false<b>``` ค่าเริ่มต้นคือ false
    ```java
    table.append("<td>font normal</td>");
    table.append("<td><b>true</b>font bold</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/pvtzqp53v/Capture.png)](https://postimg.cc/image/982ho7ac7/)
    <br><br>
    
* **ตัวอักษรตัวเอียง**<br>
ใช้ tag ```<i>true or false<i>``` ค่าเริ่มต้นคือ false
    ```java
    table.append("<td>font normal</td>");
    table.append("<td><i>true</i>font italic</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/o410vz18b/Capture.png)](https://postimg.cc/image/49ez9um0n/)
    <br><br>
    
* **ขีดเส้นใต้**<br>
ใช้ tag ```<u>รูปแบบตาม list ข้างล่าง<u>``` ค่าเริ่มต้นคือ no_underline
    - no_underline
    - single
    - single_accounting
    - double
    - double_accounting
    ```java
    table.append("<td>no underline</td>");
    table.append("<td><u>single</u>single</td>");
    table.append("<td><u>single_accounting</u>single accounting</td>");
    table.append("<td><u>double</u>double</td>");
    table.append("<td><u>double_accounting</u>double accounting</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s31.postimg.cc/h4myw66aj/Capture.png)](https://postimg.cc/image/9bwb470bb/)
    <br><br>
    
* **เปลี่ยนประเภทของคอลัมน์**<br>
ใช้ tag ```<type>รูปแบบตาม list ข้างล่าง<type>``` ค่าเริ่มต้นคือ string
    - string
    - number
    - number-money
    - number-money-float-one
    - number-money-float-two
    ```java
    table.append("<td>default string</td>");
    table.append("<td><type>number</type>7777</td>");
    table.append("<td><type>number-money</type>7777</td>");
    table.append("<td><type>number-money-float-one</type>7777</td>");
    table.append("<td><type>number-money-float-two</type>7777</td>");
    ```
    ##### ผลลัพธ์ที่ได้คือ
    [![Capture.png](https://s9.postimg.cc/mrnrdx7z3/Capture.png)](https://postimg.cc/image/pyiaxjsez/)
    <br><br>
