Attribute VB_Name = "ButtonsEvents"
Sub btn_EditProduct()
Attribute btn_EditProduct.VB_ProcData.VB_Invoke_Func = "p\n14"
ProductManager.Show
End Sub

Sub btn_EditArticles()
Attribute btn_EditArticles.VB_ProcData.VB_Invoke_Func = "k\n14"
StockManager.Show
End Sub

Sub btn_AddOrder()
Attribute btn_AddOrder.VB_ProcData.VB_Invoke_Func = "m\n14"
OrderAdder.Show
End Sub

Sub btn_AddArticle()
Attribute btn_AddArticle.VB_ProcData.VB_Invoke_Func = "n\n14"
ArticleAdder.Show
End Sub

Sub btn_ClearCommands()
Commands.Clear
End Sub

Sub btn_MergeArticles()
Attribute btn_MergeArticles.VB_ProcData.VB_Invoke_Func = "m\n14"
'ArticlesMerger.Show
End Sub

Sub btn_CreateOrder()
Attribute btn_CreateOrder.VB_ProcData.VB_Invoke_Func = "o\n14"
OrderManager.Show
End Sub

Sub InsertCmdRow_Click()
Commands.AddRow
End Sub

Sub ResetFirstsLines_Click()
Commands.ResetLines
End Sub

