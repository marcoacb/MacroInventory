# MacroInventory
Macro to automate a repetitive task.


## Description
This project is a VBA macro excel that allowed us to consolidate and track reports from our supplier UPS.


## Table of Contents
1. [Background](#background)
2. [Technologies Used](#technologies-used)
3. [Preview](#preview)
4. [Features](#features)
5. [Code Snippets](#code-snippets)
   - [VBA](#vba)
6. [Contacts and Support](#contacts-and-support)


## Background
Internal client requested to automate a repetitive task in which he consolidated two reports: the UPS report in which the entire database came and additionally extracted from another report the columns that had already been worked on the previous day to only be able to work on the new orders.

## Technologies Used
The following technologies were used to develop this dashboard:
- **VBA Excel**: For programming code.


## Preview
   - First Screen
     ![FirstScreen](Images/PrimeraPantalla.png)
     
   - Initial Form
     ![InitialForm](Images/FormularioInicial.png)
     
   - Search Folder
     ![SearchFolder](Images/BuscarCarpeta.png)

## Features
- VBA: Vba forms.


## Code Snippets

### VBA

In this section, I share some of the key methods and functions developed in VBA (Visual Basic for Applications) used throughout my projects:
- **Get the data when the file is opened**:
  ```
  Sub Actualizar()

    Dim conexion As Variant

    For Each conexion In ActiveWorkbook.Connections
        conexion.ODBCConnection.BackgroundQuery = False
    Next conexion

    ActiveWorkbook.RefreshAll

   End Sub
  ```
- **Send email when completing calculations**:
  ```
  Sub EnviarReporte_O365()

    Dim OApp As Object
    Dim OMail As Object
    Dim Body As String
    Dim RutaImagen As String
    Dim ArchivoImagen As String
    Dim NombreArchivo As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    RutaImagen = "D:\Marco - HP - Sonda\Reportes\SLA N2"
    NombreArchivo = "Reporte_BCP_SLA-L2"
    ArchivoImagen = RutaImagen & "\" & NombreArchivo & ".jpg"
        
    On Error Resume Next
    
    
    Call CrearImagenRango_O365(ThisWorkbook.Sheets("Dashboard N2"), RutaImagen, "B1:V50", NombreArchivo)
    
    Set OApp = CreateObject("Outlook.Application")
    Set OMail = OApp.CreateItem(0)
    Body = "<IMG SRC = """ & ArchivoImagen & """>"
    Fecha = Format(Now - 1, "dd/mm/yyyy")
    
    With OMail
    
        .To = Para
        .CC = CC
        .Subject = "SLA Nivel 2 - " & (Fecha)
        .BodyFormat = olFormatHTML
        .HTMLBody = Body
        .Display
        .Send
    
    End With
    
    Set OMail = Nothing
    Set OApp = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

   End Sub
  ```

## Contacts and Support
For any questions or support, contact [Marco Chang](mailto:marcochangbegazo@gmail.com).
