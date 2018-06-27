Module Sort
  Public Delegate Function Compare(ByRef v1 As Object, ByRef v2 As Object) As Integer

  Public Sub DoSort1(ByRef theData As Object, ByVal greaterThan As Compare)
    Dim outer As Integer
    Dim inner As Integer
    Dim temp As Object
    For outer = 1 To theData.count
      For inner = outer + 1 To theData.count
        If greaterThan.Invoke(theData.LoadOrder(outer), theData.LoadOrder(inner)) < 0 Then
          temp = theData.loadorder(outer)
          theData.loadorder(outer) = theData.loadorder(inner)
          theData.loadorder(inner) = temp
        End If
      Next
    Next
  End Sub

  Public Sub DoSort(ByRef theData As Object, ByVal greaterThan As Compare)
    'Dim outer As Integer
    'Dim inner As Integer
    'Dim n As Integer
    'Dim n1 As Integer
    'Dim temp As Object
    Dim nTemp As Integer
    'Dim objt As Object
    'Dim objt1 As Object


    'For n = 1 To theData.count
    '  For n1 = n + 1 To theData.count
    '    If greaterThan.Invoke(theData.loadorder(n), theData.loadorder(n1)) < 0 Then
    '      nTemp=theData.indexof(

    '      temp = theData.loadorder(n)
    '      theData.loadorder(n) = theData.loadorder(n1)
    '      theData.loadorder(n1) = temp
    '    End If
    '  Next n1
    'Next n


    For Each objt In theData
      For Each objt1 In theData
        If greaterThan.Invoke(objt, objt1) < 0 Then
          Debug.Print(objt.lecturedate & "," & objt1.lecturedate)
          'nTemp = theData.indexof(objt)
          'theData.indexof(objt) = theData.indexof(objt1)
          'theData.indexof(objt1) = nTemp

          'n = objt.loadorder
          'n1 = objt1.loadorder
          'temp = objt
          'theData.loadorder(n) = theData.loadorder(n1)
          'theData.loadorder(n1) = temp
          'objt = theData.item(n)
          'objt1 = theData.item(n)

          'temp = objt
          'objt = objt1
          'objt1 = temp
        End If
      Next
    Next
  End Sub

  Public Sub DoSort2()
    Dim date1 As Date = #8/1/2009#
    Dim date2 As Date = #8/1/2010 12:00:00 PM#
    Dim result As Integer = DateTime.Compare(date1, date2)
    Dim relationship As String

    If result < 0 Then
      relationship = "is earlier than"
    ElseIf result = 0 Then
      relationship = "is the same time as"
    Else
      relationship = "is later than"
    End If

    Console.WriteLine("{0} {1} {2}", date1, relationship, date2)
  End Sub
End Module