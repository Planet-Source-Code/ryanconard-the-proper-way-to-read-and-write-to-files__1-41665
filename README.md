<div align="center">

## The PROPER way to read and write to files


</div>

### Description

This is the PROPER, no BS way to read/write to files. Someone earlier today made a tutorial like this and it was alright but this is the easiest way. I just wanted to correct the newbies.. laterz..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RyanConard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryanconard.md)
**Level**          |Beginner
**User Rating**    |2.8 (37 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryanconard-the-proper-way-to-read-and-write-to-files__1-41665/archive/master.zip)





### Source Code

<h3>Read and Write Files</h3>
<p>
First, create a form and 2 command buttons. Name the buttons Command1 and Command2 (default names) or you change them but you must change the code as well.
<p>
Now, just copy the code below as it is into your project.
<p>
Private Sub Command1_Click()<br>
Open "C:\file.txt" For Append As #1<br>
Write #1, "Hello"<br>
Close #1<br>
End Sub
<p>
Private Sub Command2_Click()<br>
Dim Caption As String<br>
Open "C:\file.txt" For Input As #1<br>
Input #1, Caption<br>
Me.Caption = Caption<br>
Close #1<br>
End Sub
<p>
There, all set. Now just test it and see what it does. It write the world "Hello" to a file, "file.txt" and then when you click Command2 it opens that same file, takes the word out and sets it as the form caption. This can be altered to do lots of other things.
<p>
<h3>More Information</h3>
<br>
You may notice the fact that there 3 types of file procedures: Append, Output and Input. These things mean:
<p>
<b>Append:</b> This will open the selected for for Append. Append, meaning an addition to. It will open a file and write information TO it instead of deleting the contents and re-writing the data.
<br>
<b>Output:</b> This is the same as above, but this will open a file for pure Output. It will take a file and delete whatever is in it (if anything) and put only the given data into the file. Use the Output and Append correctly or you might mess something up one day.
<br>
<b>Input:</b> This opens the selected for program Input. Input, meaning it takes the data in a file and (in)puts it in one of your programs variables. Then you can use that variable to divi out the data to controls and etc.. Look at how i used it above.
<p>
Notice how these things are given numbers, such as #1. #1-#3 are what you will see most often in professional code, but all this does is give certain data thats either being inputed, outputed, or appended an identification tag so things dont get messed up. Some program may have up 10 files opened at a single time, therefore that program would have #1-#10 in it's Opened files. Think of it as a cars license plate. If every car had the same license plate number and a cop was looking for a car with that number, the cop might get the wrong person because everyones the same. So if you named everything #1 then VB would have errors in assigning correct things to do. I'm not sure, but I believe VB will only allow you to go up to #20.
<p>
<h3>Heres some examples of multitasking</h3>
<p>
Dim WinIni As String<br>
Open "C:\Windows\Win.ini" For Input As #1<br>
Open "C:\Windows\WinIniCopy.ini" For Output As #2
<br>
Open App.Path & "\log.txt" For Append As #3<br>
Input #1, WinIni<br>
Write #2, WinIni<br>
Write #3, "Today: Copied the WinIni file!"<br>
Close #1, #2, #3
<p>
Ok folks, there ya go.. Hope this helps! Now I gotta go and play GTA : Vice City. Cya! You dont have to vote.. After all, information is free!

