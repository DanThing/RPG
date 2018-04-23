Option Explicit
Option Compare Text

Imports System.IO
Imports System.Text

Namespace Global EP  

  Public Class DataFile    

    Private property get defPath() as string   
      defpath = Directory.GetCurrentDirectory()   
      console.writeline(defPath)  
    End Property
    
    Public Shared Function getDataFile(filename as string) as string   
      Try    
        Dim sr as StreamReader = new StreamReader(FS.defPath & "\data\" & filename)     
        
        Do      
          getDataFile = sr.ReadToEnd()      
          Console.WriteLine(getDataFile)     
        Loop Unitl getDataFile is Nothing     
        
        sr.Close()   
      Catch e as Exception    
        msgbox "Unable to get Data file.", VBOkOnly    
        end   
      End Try 
    End Function  
    
  End Class  
  
  Public Shared Sub Main()  
  dim gear as DataFile  
  
  gear = gear.getDataFile("gear.data") 
  
  End Sub

End Namespace
