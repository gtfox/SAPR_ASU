Option Explicit

Private Sub UserForm_Activate()
ProgressBar1.Value = 50
  WB1.Navigate "https://avselectro.ru/search/index.php?q=MVA20-1-032-C"
End Sub