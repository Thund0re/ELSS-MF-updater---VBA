Attribute VB_Name = "mdlWebscraper"




Sub Get_NAV()

Dim IE

Dim NAV, Units As Double

'Creating Internet Explorer instance
Set IE = CreateObject("InternetExplorer.Application")
'Navigate to get MF's NAV URL
IE.navigate "https://www.etmoney.com/mutual-funds/axis-long-term-equity-fund-growth/10480"

'IE.Visible = True

Do
    DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

'Accesss HTML DOM to get NAV
Set Doc = IE.document
NAV = Doc.getElementsByClassName("amount")(0).innerText


Sheets("AXISMF").Range("I3") = Date
Sheets("AXISMF").Range("I3").NumberFormat = "dd-mmmm-yyyy"

Sheets("AXISMF").Range("I5") = Right(Trim(NAV), 5)

NAV = Sheets("AXISMF").Range("I5")

Units = Sheets("AXISMF").Range("E5")

Sheets("AXISMF").Range("J5") = NAV * Units

'Close IE and free up memory
IE.Quit
Set IE = Nothing

MsgBox "Updated"



End Sub

