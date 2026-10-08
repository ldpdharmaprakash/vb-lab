**P R A C T I C A L L A B G U I D E** 

# **Visual Basic 6.0** 

Exercises 1 to 7 

_Step-by-step: controls  →  properties  →  code  →  output_ 

##### **S T U D E N T E D I T I O N** 

Exercises covered in this guide 

- **1** Arithmetic Calculator 

- **2** Fibonacci Series and Sum of N Numbers 

- **3** Menu Driven MDI Application 

- **4** Rich Text Box Control 

- **5** Common Dialog Control 

- **6** Animation Using Timers 

- **7** Number Conversion 

**Visual Basic 6.0** 

Exercises 1 – 7 

## **How to Use This Guide** 

Every exercise follows the **same four steps** . Follow them in order and your program will work the first time. 

|**Step**|**What you do**|**Where in VB6**|
|---|---|---|
|**1  Controls**|Draw the controls on the form|Toolbox (left side)|
|**2  Properties**|Set<br>(Name),<br>Captionand other properties|Properties window (right side,**F4**)|
|**3  Code**|Double-click a control and type/paste the<br>code|Code window (**F7**)|
|**4  Run & Test**|Run the program and compare with the<br>sample output|**F5**or the Start button|



#### **Golden rules** 

- Give controls the **exact names** shown in the guide. The code uses these names. 

- Always put <mark>Option Explicit</mark> at the top. It catches spelling mistakes in variable names. 

- Save often: **File → Save Project** (Ctrl + S). Keep the <mark>.frm</mark> and <mark>.vbp</mark> files in one folder. 

- If a control is missing from the toolbox, add it from **Project → Components** . 

#### **Colour guide used in this document** 

|**Box**|**Meaning**|
|---|---|
|**AIM**(blue)|What the exercise teaches|
|**EXAM / VIVA TIP**(gold)|Points to remember for exams and viva|
|**LAB APPLICATION**(green)|Setup actions to do in the VB6 environment|
|**COMMON MISTAKE**(red)|Typical errors and how to avoid them|



#### **Lab index** 

|**No.**|**Exercise**|**Main controls / concept**|
|---|---|---|
|1|Arithmetic Calculator|TextBox, CommandButton,<br>InStr,<br>Val|
|2|Fibonacci and Sum of N|For … Next,<br>If … Then|
|3|Menu Driven MDI Application|MDIForm, Menu Editor,<br>Arrange|
|4|Rich Text Box Control|Drive/Dir/FileListBox, RichTextBox|
|5|Common Dialog Control|CommonDialog: Open, Save, Print, Font, Color|
|6|Animation Using Timers|Timer, Image,<br>Left/<br>Top|
|7|Number Conversion|Hex,<br>Oct, user-defined function|



Page 2 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 1   Arithmetic Calculator** 

###### **AIM** 

To construct a simple arithmetic calculator in Visual Basic 6.0 that performs addition, subtraction, multiplication and division. 

**Concept:** Every button adds its character to a text variable <mark>temp</mark> . When **=** is pressed, the program finds the operator with <mark>InStr</mark> , splits the text into two numbers using <mark>Left</mark> and <mark>Mid</mark> , and calculates the result. 

#### **STEP 1  Create the Controls** 

**IMPORTANT – NO CONTROL ARRAYS** 

Create every button **individually** <mark>(Command1</mark> , <mark>Command2</mark> … <mark>Command17</mark> ). Do **not** copy-paste a button and answer **Yes** to “Do you want to create a control array?”. A control array adds an <mark>Index</mark> parameter and the code below will not work. 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|TextBox|Text1|Display screen|
|CommandButton|Command1|1|
|CommandButton|Command2|2|
|CommandButton|Command3|3|
|CommandButton|Command4|4|
|CommandButton|Command5|5|
|CommandButton|Command6|6|
|CommandButton|Command7|7|
|CommandButton|Command8|8|
|CommandButton|Command9|9|
|CommandButton|Command10|0|
|CommandButton|Command11|.|
|CommandButton|Command12|+|
|CommandButton|Command13|-|
|CommandButton|Command14|*|
|CommandButton|Command15|/|
|CommandButton|Command16|=|
|CommandButton|Command17|AC|
|CommandButton|Command18|Exit|



Page 3 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **Form layout (sketch)** 

┌────────────────────────┐ <mark>Text1         │ │ └────────────────────────┘ [ 1 ]  [ 2 ]  [ 3 ]  [ + ] [ 4 ]  [ 5 ]  [ 6 ]  [ - ] [ 7 ]  [ 8 ]  [ 9 ]  [ * ] [ 0 ]  [ . ]  [AC ]  [ / ]</mark> [Exit]        [ = ] 

#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form|(Name)|Form1|
|Form|Caption|Arithmetic Calculator|
|Form|StartUpPosition|2 - CenterScreen|
|Text1|(Name)|Text1|
|Text1|Text|0|
|Text1|Alignment|1 - Right Justify|



#### **STEP 3  Write the Code** 

**◆ CODE  ·  Form1** <mark>Option Explicit Dim a As String Dim b As Double Dim temp As String</mark> _<mark>' holds the expression typed so far, e.g. 12+56</mark>_ <mark>Private Sub Form_Load() Text1.Text = "0" temp = "" b = 0 a = "" End Sub Private Sub Command1_Click() temp = temp & "1" Text1.Text = temp End Sub Private Sub Command2_Click() temp = temp & "2" Text1.Text = temp End Sub</mark> 

Page 4 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

<mark>Private Sub Command3_Click() temp = temp & "3" Text1.Text = temp End Sub Private Sub Command4_Click() temp = temp & "4" Text1.Text = temp End Sub Private Sub Command5_Click() temp = temp & "5" Text1.Text = temp End Sub Private Sub Command6_Click() temp = temp & "6" Text1.Text = temp End Sub Private Sub Command7_Click() temp = temp & "7" Text1.Text = temp End Sub Private Sub Command8_Click() temp = temp & "8" Text1.Text = temp End Sub Private Sub Command9_Click() temp = temp & "9" Text1.Text = temp End Sub Private Sub Command10_Click() temp = temp & "0" Text1.Text = temp End Sub Private Sub Command11_Click()</mark> _<mark>' add a decimal point only once</mark>_ <mark>If InStr(temp, ".") = 0 Then If temp = "" Then temp = "0." Else temp = temp & "." End If End If Text1.Text = temp End Sub Private Sub Command12_Click() temp = temp & "+" Text1.Text = temp End Sub</mark> 

Page 5 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

<mark>Private Sub Command13_Click() temp = temp & "-" Text1.Text = temp End Sub Private Sub Command14_Click() temp = temp & "*" Text1.Text = temp End Sub Private Sub Command15_Click() temp = temp & "/" Text1.Text = temp End Sub Private Sub Command16_Click() Dim position As Integer Dim f</mark> i <mark>rstNumber As Double Dim secondNumber As Double Dim result As Double position = InStr(temp, "+")</mark> _<mark>' where is the operator?</mark>_ <mark>If position > 0 Then frstNumber = Val(Left(temp, position -</mark> i <mark>1)) secondNumber = Val(Mid(temp, position + 1)) result = f</mark> i <mark>rstNumber + secondNumber Text1.Text = result temp = CStr(result) Exit Sub End If position = InStr(temp, "-")</mark> _<mark>' where is the operator?</mark>_ <mark>If position > 0 Then frstNumber = Val(Left(temp, position -</mark> i <mark>1)) secondNumber = Val(Mid(temp, position + 1)) result = f</mark> i <mark>rstNumber - secondNumber Text1.Text = result temp = CStr(result) Exit Sub End If position = InStr(temp, "*")</mark> _<mark>' where is the operator?</mark>_ <mark>If position > 0 Then frstNumber = Val(Left(temp, position -</mark> i <mark>1)) secondNumber = Val(Mid(temp, position + 1)) result = f</mark> i <mark>rstNumber * secondNumber Text1.Text = result temp = CStr(result) Exit Sub End If position = InStr(temp, "/") If position > 0 Then frstNumber = Val(Left(temp, position -</mark> i <mark>1)) secondNumber = Val(Mid(temp, position + 1))</mark> 

<mark>If secondNumber = 0 Then</mark> _<mark>' avoid division by zero</mark>_ 

Page 6 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

<mark>Text1.Text = "Error" temp = "" Exit Sub End If result = f</mark> i <mark>rstNumber / secondNumber Text1.Text = result temp = CStr(result) Exit Sub End If End Sub Private Sub Command17_Click() temp = "" a = "" b = 0 Text1.Text = "0"</mark> End Sub <mark>Private Sub Command18_Click() End End Sub</mark> 

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Click the buttons and watch the expression build up in the display. The expression stays **visible** until you press **=** . 

|**Button pressed**|**Display shows**|
|---|---|
|1|1|
|2|12|
|+|12+|
|5|12+5|
|6|12+56|
|=|68|



Page 7 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 



<!-- Start of picture text -->
Pepe fief -] fe<br><!-- End of picture text -->



<!-- Start of picture text -->
oftied el =]<br>otetiedel<br>fiefs}x]<br>Edel<br><!-- End of picture text -->

Page 8 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 2   Fibonacci Series and Sum of N Numbers** 

###### **AIM** 

To use loops ( <mark>For … Next</mark> ) and decision making ( <mark>If … Then</mark> ) to (A) generate the Fibonacci series and (B) find the sum of the first N numbers. 

One form is used for both parts: **2(A) Fibonacci Series** and **2(B) Sum of N Numbers** . 

#### **STEP 1  Create the Controls** 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|Label|Label1|Enter N:|
|TextBox|Text1|Input – value of N|
|Label|Label2|Result:|
|TextBox|Text2|Output – answer|
|CommandButton|Command1|Sum of N Numbers|
|CommandButton|Command2|Fibonacci|
|CommandButton|Command3|Clear  (optional)|
|CommandButton|Command4|Exit  (optional)|



#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form|(Name)|Form1|
|Form|Caption|Loops and Decision Making|
|Text1|Text|(empty)|
|Text2|Text|(empty)|
|Text2|MultiLine|True|



Page 9 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 3  Write the Code** 

**◆ CODE  ·  Form1** <mark>Option Explicit Dim n As Long Dim i As Long Dim sum As Long Dim a As Long Dim b As Long Dim f</mark> i <mark>b As Long Dim result As String</mark> _<mark>' ---------- 2(B) Sum of N numbers ----------</mark>_ <mark>Private Sub Command1_Click() n = Val(Text1.Text) If n <= 0 Then MsgBox "Enter a positive number", vbExclamation, "Error" Text1.SetFocus Exit Sub End If sum = 0 For i = 1 To n sum = sum + i Next i Text2.Text = sum End Sub</mark> _<mark>' ---------- 2(A) Fibonacci series ----------</mark>_ <mark>Private Sub Command2_Click() n = Val(Text1.Text) If n <= 0 Then MsgBox "Enter a positive number", vbExclamation, "Error" Text1.SetFocus Exit Sub End If a = 0 b = 1 result = "" For i = 1 To n result = result & a If i < n Then result = result & " " End If fb = a + b</mark> i _<mark>' next term = sum of previous two</mark>_ <mark>a = b b = f</mark> i <mark>b Next i Text2.Text = result End Sub Private Sub Command3_Click() Text1.Text = ""</mark> 

Page 10 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

<mark>Text2.Text = "" Text1.SetFocus End Sub Private Sub Command4_Click() End</mark> End Sub 

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Type a value in **Text1** and click a button. Example with **N = 5** : 

|**Button**|**Calculation**|**Output in Text2**|
|---|---|---|
|Sum of N Numbers|1 + 2 + 3 + 4 + 5|15|
|Fibonacci|first 5 terms|0 1 1 2 3|



Page 11 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 3   Menu Driven MDI Application** 

###### **AIM** 

To create a Multiple Document Interface (MDI) application with a menu bar that opens child forms and arranges them in different layouts. 

**Concept:** An **MDI Form** is the parent window. Other forms whose <mark>MDIChild</mark> property is **True** open *inside* it. A **menu** lets the user choose which child form to open. 

#### **STEP 1  Create the Controls** 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|MDI Form|MDIForm1|Parent window  (Project → Add MDI Form)|
|Form|Form1|Child form 1|
|Form|Form2|Child form 2  (Project → Add Form)|



#### **Menu structure** 

Click **MDIForm1** , then open **Tools → Menu Editor** (shortcut **Ctrl + E** ) and create: 

File <mark>├── Form 1 └── Form 2 Window ├── Cascade ├── Tile Horizontal</mark> └── Tile Vertical 

|**Caption**|**(Name)**|**Indent**|
|---|---|---|
|&File|file|0 (main menu)|
|Form 1|f1|1 (one arrow →)|
|Form 2|f2|1|
|&Window|window|0 (main menu)|
|Cascade|cas|1|
|Tile Horizontal|hari|1|
|Tile Vertical|verti|1|



Page 12 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form1|MDIChild|True|
|Form2|MDIChild|True|
|Form1 / Form2|Caption|Form 1  /  Form 2  (any suitable text)|



###### **PROJECT SETTING** 

Go to **Project → Project1 Properties → General** and set **Startup Object = MDIForm1** . If you forget this, the program will start with a child form and give an error. 

#### **Final project structure** 

Project1 

- <mark>├── MDIForm1   (parent, contains the menu) ├── Form1      (MDIChild = True)</mark> └── Form2      (MDIChild = True) 

#### **STEP 3  Write the Code** 

**◆ CODE  ·  MDIForm1** <mark>Private Sub f1_Click() Form1.Show</mark> _<mark>' open child form 1</mark>_ <mark>End Sub Private Sub f2_Click() Form2.Show</mark> _<mark>' open child form 2</mark>_ <mark>End Sub Private Sub cas_Click() MDIForm1.Arrange vbCascade End Sub Private Sub hari_Click() MDIForm1.Arrange vbTileHorizontal End Sub Private Sub verti_Click() MDIForm1.Arrange vbTileVertical</mark> End Sub 

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Use the menus to test the application: 

■ **File → Form 1** and **File → Form 2** open the two child forms inside the MDI window. 

Page 13 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

- **Window → Cascade** stacks the windows one over another. 

- **Window → Tile Horizontal** places the windows one above the other. 

- **Window → Tile Vertical** places the windows side by side. 

Page 14 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 4   Rich Text Box Control** 

###### **AIM** 

To open, edit and save a text file using the DriveListBox, DirListBox, FileListBox and RichTextBox controls. 

**Concept:** The three list boxes work together as a file browser. Choosing a drive updates the folder list, choosing a folder updates the file list, and clicking a file loads it into the RichTextBox. 

#### **STEP 1  Create the Controls** 

**ADD THE RICHTEXTBOX TO THE TOOLBOX** 

If the RichTextBox is not visible in the toolbox: 

**Project → Components → Microsoft Rich Textbox Control 6.0 → Apply → OK** 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|DriveListBox|Drive1|Select the drive|
|DirListBox|Dir1|Select the folder|
|FileListBox|File1|Shows the files|
|TextBox|Text1|Shows the selected file path|
|RichTextBox|RichTextBox1|Displays and edits the file|
|CommandButton|Command1|Open|
|CommandButton|Command2|Save|
|CommandButton|Command3|Exit|



#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form|Caption|Rich Text Box File Editor|
|File1|Pattern|*.txt|
|RichTextBox1|MultiLine|True|



Page 15 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 3  Write the Code** 

|**CODE  ·  Form1**<br>**◆**<br>Option Explicit|
|---|
|Private SubForm_Load()<br> Drive1.Drive="D:"|
|Dir1.Path ="D:\BCAS\RDBMS\vb lab" _' change to your own folder_<br> File1.Path =Dir1.Path<br> File1.Pattern ="*.txt"<br>End Sub|
|Private SubDrive1_Change()<br> Dir1.Path =Drive1.Drive _' drive changed -> update folders_<br>End Sub|
|Private SubDir1_Change()<br> File1.Path =Dir1.Path_' folder changed -> update files_<br> File1.Pattern ="*.txt"<br>End Sub<br>Private SubFile1_Click()<br> Text1.Text=File1.Path &"\"&File1.FileName<br> RichTextBox1.LoadFileText1.Text,rtfText<br>End Sub|
|Private SubCommand1_Click()<br> If Text1.Text="" Then<br>MsgBox"Please select a text file"<br> Exit Sub<br> End If<br> RichTextBox1.LoadFileText1.Text,rtfText<br>End Sub|
|Private SubCommand2_Click()<br> If Text1.Text="" Then<br>MsgBox"Please select a text file"<br> Exit Sub<br> End If<br> RichTextBox1.SaveFileText1.Text,rtfText<br>MsgBox"FILE SAVED"<br>End Sub<br>Private SubCommand3_Click()<br> End<br>End Sub|



###### **HOW THE FILE PATH IS BUILT** 

If <mark>File1.Path</mark> is <mark>D:\BCAS\RDBMS\vb lab</mark> and <mark>File1.FileName</mark> is <mark>sample.txt</mark> , then <mark>File1.Path & "\" & File1.FileName</mark> gives **D:\BCAS\RDBMS\vb lab\sample.txt** 

Make sure the folder in <mark>Form_Load</mark> exists on your computer and contains at least one <mark>.txt</mark> file. 

Page 16 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Select a drive, a folder and a file. The file path appears in **Text1** and the content appears in the **RichTextBox** . Edit the text and click **Save** . 

Page 17 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 5   Common Dialog Control** 

###### **AIM** 

To display the standard Windows dialog boxes – Open, Save, Printer, Font and Color – using the Microsoft Common Dialog Control. 

#### **STEP 1  Create the Controls** 

###### **ADD THE COMMON DIALOG CONTROL** 

**Project → Components → Microsoft Common Dialog Control 6.0 → Apply → OK** , then draw it on the form. It is a **non-visual control** – it will not appear when the program runs. 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|CommandButton|Command1|Save|
|CommandButton|Command2|Open|
|CommandButton|Command3|Print|
|CommandButton|Command4|Font|
|CommandButton|Command5|Color|
|CommandButton|Command6|Exit|
|CommonDialog|CommonDialog1|Provides the dialog boxes|



#### **Form layout (sketch)** 

COMMON DIALOG CONTROL 

<mark>[   Save   ] [   Open   ] [   Print  ] [   Font   ] [   Color  ] [   Exit   ]</mark> CommonDialog1  (non-visual) 

#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form|Caption|Common Dialog Control|
|CommonDialog1|(Name)|CommonDialog1|



Page 18 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 3  Write the Code** 



<!-- Start of picture text -->
◆  CODE  ·  Form1<br>Option Explicit<br>Private Sub Command1_Click()<br>CommonDialog1.ShowSave<br>End Sub<br>Private Sub Command2_Click()<br>CommonDialog1.ShowOpen<br>End Sub<br>Private Sub Command3_Click()<br>CommonDialog1.ShowPrinter<br>End Sub<br>Private Sub Command4_Click()<br>CommonDialog1.ShowFont<br>End Sub<br>Private Sub Command5_Click()<br>CommonDialog1.ShowColor<br>End Sub<br>Private Sub Command6_Click()<br>End<br>End Sub<br><!-- End of picture text -->

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Click each button and the matching standard Windows dialog box opens. 

Page 19 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 6   Animation Using Timers** 

###### **AIM** 

To create a simple animation by moving an Image control across the form using the Timer control. 

**Concept:** A Timer fires its <mark>Timer</mark> event again and again after a fixed <mark>Interval</mark> (in milliseconds). Each time it fires, the image is moved a little, which looks like animation. 

#### **STEP 1  Create the Controls** 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|Image|Image1|Picture that will move|
|Timer|Timer1|Repeats the movement|
|CommandButton|Command1|START|
|CommandButton|Command2|STOP|
|CommandButton|Command3|EXIT|



Draw the **Image** control on the form and load any <mark>.bmp</mark> , <mark>.jpg</mark> or <mark>.gif</mark> file using its **Picture** property. 

#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Image1|(Name)|Image1|
|Image1|Picture|(choose any picture)|
|Timer1|Enabled|False|
|Timer1|Interval|100|



Page 20 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 3  Write the Code** 

|**CODE  ·  Form1**<br>**◆**<br>Option Explicit|
|---|
|Private SubForm_Load()<br>|
|Timer1.Enabled =False<br> Timer1.Interval =100 _' 100 ms = 10 moves per second_<br>|
|End Sub<br>Private SubCommand1_Click()|
|Timer1.Enabled =True _' START_<br>End Sub|
|Private SubCommand2_Click()<br> Timer1.Enabled =False _' STOP_<br>End Sub|
|Private SubCommand3_Click()<br> End _' EXIT_<br>|
|End Sub|
|Private SubTimer1_Timer()|
|Image1.Left =Image1.Left +10 _' move right_|
|If Image1.Left >=2600 Then|
|Image1.Top =Image1.Top -2 _' then drift upward_<br> End If<br>End Sub|



#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

- Click **START** – the timer is enabled and the image moves to the right. 

- Click **STOP** – the timer is disabled and the image halts. 

- Click **EXIT** – the application closes. 

Page 21 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

### **EXERCISE 7   Number Conversion** 

###### **AIM** 

To convert a decimal number into its Binary, Octal and Hexadecimal equivalents. 

**Concept:** VB6 has built-in functions <mark>Hex()</mark> and <mark>Oct()</mark> . There is no built-in binary function, so we write our own function <mark>dec()</mark> that repeatedly divides the number by 2 and collects the remainders. 

#### **STEP 1  Create the Controls** 

|**Control**|**(Name)**|**Caption / Purpose**|
|---|---|---|
|Label|Label1|Enter Decimal Number|
|TextBox|Text1|Input|
|Label|Label2|Result|
|TextBox|Text2|Output|
|CommandButton|Command1|Binary|
|CommandButton|Command2|Hexadecimal|
|CommandButton|Command3|Octal|
|CommandButton|Command4|Clear|
|CommandButton|Command5|Exit|



#### **Form layout (sketch)** 



<!-- Start of picture text -->
        NUMBER CONVERSION<br>        Enter Decimal Number<br>┌────────────────────┐<br>               Text1        │ │<br>└────────────────────┘<br>        [ Binary ]  [ Hex ]  [ Octal ]<br>        Result<br>┌────────────────────┐<br>               Text2        │ │<br>└────────────────────┘<br>        [ Clear ]   [ Exit ]<br><!-- End of picture text -->

Page 22 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

#### **STEP 2  Set the Properties** 

|**Object**|**Property**|**Value**|
|---|---|---|
|Form|Caption|Number Conversion|
|Text1|Text|(empty)|
|Text2|Text|(empty)|



#### **STEP 3  Write the Code** 

**◆ CODE  ·  Form1** <mark>Option Explicit</mark> _<mark>' Converts a decimal number to binary text</mark>_ <mark>Function dec(ByVal g As Long) As String Dim s As String If g = 0 Then dec = "0" Exit Function End If Do Until g = 0 If (g Mod 2) Then</mark> _<mark>' remainder 1</mark>_ <mark>s = "1" & s Else</mark> _<mark>' remainder 0</mark>_ <mark>s = "0" & s End If g = g \ 2</mark> _<mark>' integer division</mark>_ <mark>Loop dec = s End Function Private Sub Command1_Click() Text2.Text = dec(Val(Text1.Text)) End Sub Private Sub Command2_Click() Text2.Text = Hex(Val(Text1.Text)) End Sub Private Sub Command3_Click() Text2.Text = Oct(Val(Text1.Text)) End Sub Private Sub Command4_Click() Text1.Text = "" Text2.Text = "" Text1.SetFocus End Sub</mark> 

Page 23 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

<mark>Private Sub Command5_Click() End</mark> End Sub 

#### **STEP 4  Run and Test** 

Press **F5** (or click the **Start** button) to run the project. 

Enter a whole number, for example **25** , and click each button: 

|**Button**|**Output for 25**|
|---|---|
|Binary|11001|
|Octal|31|
|Hexadecimal|19|



###### **EXAM / VIVA TIP** 

Use <mark>g \ 2</mark> (integer division) instead of <mark>g / 2</mark> because g is a <mark>Long</mark> . The binary digits are collected **right to left** , so each new digit is added at the **front** : <mark>s = "1" & s</mark> . 

###### **HOW 25 BECOMES BINARY** 

25 ÷ 2 = 12 remainder **1** →   12 ÷ 2 = 6 remainder **0** →   6 ÷ 2 = 3 remainder **0** 3 ÷ 2 = 1 remainder **1** →   1 ÷ 2 = 0 remainder **1** 

Read the remainders from bottom to top: **11001** 

Page 24 of 25 

**Visual Basic 6.0** 

Exercises 1 – 7 

## **Quick Revision Checklist** 

Use this page to check that you have placed all the required controls before writing the code. 

|**Ex.**|**Controls required**|**Key names**|
|---|---|---|
|1|1 TextBox + 17 CommandButtons|Text1,<br>Command1–<br>Command17(no control<br>array)|
|2|2 Labels, 2 TextBoxes, 4 CommandButtons|Text1input,<br>Text2result,<br>Command1Sum,<br>Command2Fibonacci|
|3|1 MDI Form, 2 child forms, menu|MDIForm1,<br>Form1,<br>Form2(MDIChild = True)|
|4|Drive, Dir, File list boxes, TextBox,<br>RichTextBox, 3 buttons|Drive1,<br>Dir1,<br>File1,<br>Text1,<br>RichTextBox1|
|5|CommonDialog + 6 CommandButtons|CommonDialog1,<br>Command1–<br>Command6|
|6|Image, Timer, 3 CommandButtons|Image1,<br>Timer1, START / STOP / EXIT|
|7|2 Labels, 2 TextBoxes, 5 CommandButtons|Text1input,<br>Text2output,<br>Command1–<br>Command5|



#### **Common errors and their fixes** 

|**Ex.**|**Problem**|**Correct version**|
|---|---|---|
|1|Command2_Click(Index As Integer)– button<br>was created as a control array|Private Sub Command2_Click()– create<br>buttons individually|
|1|Expression disappears while typing|Keep the whole expression in<br>tempand show it<br>in<br>Text1|
|2|Stray text such as<br>a = b xyz|a = b|
|3|vbtilecascade|vbCascade|
|4|File2.Pattern = "*.txt"|File1.Pattern = "*.txt"|
|4|Fixed path<br>C:/SAMPLE/hello1.txt|File1.Path & "\" & File1.FileName|
|5|Command7_Clickcalls undefined<br>editer|Remove it, or write the<br>editerprocedure first|
|7|g = g / 2with a<br>Longvariable|g = g \ 2|
|7|Function name used on both sides of an<br>assignment|Build the answer in a local variable<br>s, then<br>dec =<br>s|



###### **FINAL TIP** 

Teaching routine for every program: **create the components → set `(Name)` and `Caption` → paste the code → run and test the output.** 

Page 25 of 25 

