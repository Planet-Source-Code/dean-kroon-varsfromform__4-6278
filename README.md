<div align="center">

## VarsFromForm


</div>

### Description

This subroutine creates a variable (on the fly) for each form item found on Request.Forms or Request.QueryString collections. The variable is given the same name as the form items name. It was built to duplicate some of PHPs built in functionality.
 
### More Info
 


Ex.

' this form lives in some asp or html page

<form method="post" action="doit.asp">

<input type="text" name="name">

<input type="text" name="Address1">

<input type="text" name="Address2">

</form>

' the user enters:

Dean Kroon

421 Blue Earth St

PO Box 123

' this code lives in doit.asp

<%

Call VarsFromForm

Response.write name & "<br>" & address1 & ", " & address2

%>

' The output looks like

Dean Kroon

421 Blue Earth St, PO Box 123


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dean Kroon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dean-kroon.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dean-kroon-varsfromform__4-6278/archive/master.zip)





### Source Code

```
public sub VarsFromForm
	for each item in request.form
		execute(item & "=""" & Replace(request.form(item), Chr(34), Chr(34) & Chr(34)) & """")
	next
	for each item in request.QueryString
		execute(item & "=""" & Replace(request.QueryString(item), Chr(34), Chr(34) & Chr(34)) & """")
	next
end sub
```

