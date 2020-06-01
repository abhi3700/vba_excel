* `=` 
	- Equal Sign Sets a Property, Object, or Variable in VBA
	- E.g-1: set the Value property of Cell A1 to the word hello.
```vbs
Range("A1").Value = "Hello"
```
	- E.g-2: set the value of a variable
```vbs
lRow = 10
```
	- E.g-3: The equal sign can also be used in If statements as a comparison operator to evaluate a condition.
```vbs
If (4+4) = 9 Then "All the textbooks in the world must be re-printed!"
```

* `:=` 
	- The Colon Equal Sign Sets a Value of a Parameter for a Property or Method
	- E.g-1: The Worksheets.Add method has four optional parameters. as 
```md
Worksheets.Add ([Before], [After], [Count], [Type])
```
	
	The following line of code will add a worksheet after the active sheet.
```vbs
Worksheets.Add After:=Activesheet
```