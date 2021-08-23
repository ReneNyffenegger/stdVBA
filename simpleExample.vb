option explicit

sub Main()
  'Create an array
  Dim arr as stdArray
  set arr = stdArray.Create(1,2,3,4,5,6,7,8,9,10) 'Can also call CreateFromArray

  'Demonstrating join, join will be used in most of the below functions
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10
  Debug.Print arr.join("|")                                              '1|2|3|4|5|6|7|8|9|10

  'Basic operations
  arr.push 3
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10,3
  Debug.Print arr.pop()                                                  '3
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10
  Debug.Print arr.concat(stdArray.Create(11,12,13)).join                 '1,2,3,4,5,6,7,8,9,10,11,12,13
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10 'concat doesn't mutate object
  Debug.Print arr.includes(3)                                            'True
  Debug.Print arr.includes(34)                                           'False

  'More advanced behaviour when including callbacks! And VBA Lamdas!!
  Debug.Print arr.Map(stdLambda.Create("$1+1")).join          '2,3,4,5,6,7,8,9,10,11
  Debug.Print arr.Reduce(stdLambda.Create("$1+$2"))           '55 ' I.E. Calculate the sum
  Debug.Print arr.Reduce(stdLambda.Create("application.worksheetFunction.Max($1,$2)"))      '10 ' I.E. Calculate the maximum
  Debug.Print arr.Filter(stdLambda.Create("$1>=5")).join      '5,6,7,8,9,10
  
  'Execute property accessors with Lambda syntax
' Debug.Print arr.Map(stdLambda.Create("ThisWorkbook.Sheets($1)")) _ 
'                .Map(stdLambda.Create("$1.Name")).join(",")            'Sheet1,Sheet2,Sheet3,...,Sheet10
  
  'Execute methods with lambdas and enumerate over enumeratable collections:
' Call stdEnumerator.CreateFromArray(Application.Workbooks).forEach(stdLambda.Create("$1.Save"))
  
  'We even have if statement!
  With stdLambda.Create("if $1 then ""lisa"" else ""bart""")
    Debug.Print .Run(true)                                              'lisa
    Debug.Print .Run(false)                                             'bart
  End With
  
  'Execute custom functions
' Debug.Print arr.Map(stdCallback.CreateFromModule("ModuleMain","CalcArea")).join  '3.14159,12.56636,28.274309999999996,50.26544,78.53975,113.09723999999999,153.93791,201.06176,254.46879,314.159

  'Let's move onto regex:
  Dim oRegex as stdRegex
  set oRegex = stdRegex.Create("(?<county>[A-Z])-(?<city>\d+)-(?<street>\d+)","i")

  Dim oRegResult as object
  set oRegResult = oRegex.Match("D-040-1425")
  Debug.Print oRegResult("county") 'D
  Debug.Print oRegResult("city")   '040
  
  'And getting all the matches....
  Dim sHaystack as string: sHaystack = "D-040-1425;D-029-0055;A-100-1351"
' Debug.Print stdEnumerator.CreateFromEnumVARIANT(oRegex.MatchAll(sHaystack)).map(stdLambda.Create("$1.item(""county"")")).join 'D,D,A
  Debug.Print stdEnumerator.CreateFromIEnumVARIANT(oRegex.MatchAll(sHaystack)).map(stdLambda.Create("$1.item(""county"")")).join 'D,D,A
  
  'Dump regex matches to range:
  '   D,040,040-1425
  '   D,029,029-0055
  '   A,100,100-1351
  Range("A3:C6").value = oRegex.ListArr(sHaystack, Array("$county","$city","$city-$street"))
  
  'Copy some data to the clipboard:
  Range("A1").value = "Hello there"
  Range("A1").copy
  Debug.Print stdClipboard.Text 'Hello there
  stdClipboard.Text = "Hello world"
  Debug.Print stdClipboard.Text 'Hello world

  'Copy files to the clipboard.
  Dim files as collection
  set files = new collection
  files.add "C:\File1.txt"
  files.add "C:\File2.txt"
  set stdClipboard.files = files

  'Save a chart as a file
' Sheets("Sheet1").ChartObjects(1).copy
' Call stdClipboard.Picture.saveAsFile("C:\test.bmp",false,null) 'Use IPicture interface to save to disk as image
End Sub

Public Function CalcArea(ByVal radius as Double) as Double
  CalcArea = 3.14159*radius*radius
End Function
