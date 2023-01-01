<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ФайлToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПерезагрузкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбновитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИнструментыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтчётПоВыданнойТехникиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЖурналToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправочникиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПолучателиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТехникаToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыдатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДобавитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОПрограммеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОАвтореToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВидToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.БольшиеИконкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДеталиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.МаленькиеИконкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СписокToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НазваниеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ВыдатьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ЖурналВыдачДанногоУстройстваToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТекущийОтчетДанногоУстройстваToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ВидToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.БольшиеИконкиToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДеталиToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.МаленькиеИконкиToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.СписокToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.НазваниеToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбновитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДобавитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.AllowColumnReorder = True
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HideSelection = False
        Me.ListView1.LargeImageList = Me.ImageList2
        Me.ListView1.Location = New System.Drawing.Point(0, 24)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(523, 315)
        Me.ListView1.SmallImageList = Me.ImageList1
        Me.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "камера64.ico")
        Me.ImageList2.Images.SetKeyName(1, "проектор64.ico")
        Me.ImageList2.Images.SetKeyName(2, "Phone64.ico")
        Me.ImageList2.Images.SetKeyName(3, "ноутбук64.ico")
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "камера16.ico")
        Me.ImageList1.Images.SetKeyName(1, "проектор16.ico")
        Me.ImageList1.Images.SetKeyName(2, "Phone16.ico")
        Me.ImageList1.Images.SetKeyName(3, "ноутбук16.ico")
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ФайлToolStripMenuItem, Me.ИнструментыToolStripMenuItem, Me.СправочникиToolStripMenuItem, Me.СправкаToolStripMenuItem, Me.ВидToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(523, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ФайлToolStripMenuItem
        '
        Me.ФайлToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПерезагрузкаToolStripMenuItem, Me.ОбновитьToolStripMenuItem, Me.ToolStripMenuItem1, Me.ВыходToolStripMenuItem})
        Me.ФайлToolStripMenuItem.Name = "ФайлToolStripMenuItem"
        Me.ФайлToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ФайлToolStripMenuItem.Text = "Файл"
        '
        'ПерезагрузкаToolStripMenuItem
        '
        Me.ПерезагрузкаToolStripMenuItem.Name = "ПерезагрузкаToolStripMenuItem"
        Me.ПерезагрузкаToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ПерезагрузкаToolStripMenuItem.Text = "Перезагрузка"
        '
        'ОбновитьToolStripMenuItem
        '
        Me.ОбновитьToolStripMenuItem.Name = "ОбновитьToolStripMenuItem"
        Me.ОбновитьToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ОбновитьToolStripMenuItem.Text = "Обновить"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(145, 6)
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'ИнструментыToolStripMenuItem
        '
        Me.ИнструментыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОтчётПоВыданнойТехникиToolStripMenuItem, Me.ЖурналToolStripMenuItem})
        Me.ИнструментыToolStripMenuItem.Name = "ИнструментыToolStripMenuItem"
        Me.ИнструментыToolStripMenuItem.Size = New System.Drawing.Size(95, 20)
        Me.ИнструментыToolStripMenuItem.Text = "Инструменты"
        '
        'ОтчётПоВыданнойТехникиToolStripMenuItem
        '
        Me.ОтчётПоВыданнойТехникиToolStripMenuItem.Name = "ОтчётПоВыданнойТехникиToolStripMenuItem"
        Me.ОтчётПоВыданнойТехникиToolStripMenuItem.Size = New System.Drawing.Size(236, 22)
        Me.ОтчётПоВыданнойТехникиToolStripMenuItem.Text = "Отчёт по выданной технике..."
        '
        'ЖурналToolStripMenuItem
        '
        Me.ЖурналToolStripMenuItem.Name = "ЖурналToolStripMenuItem"
        Me.ЖурналToolStripMenuItem.Size = New System.Drawing.Size(236, 22)
        Me.ЖурналToolStripMenuItem.Text = "Журнал..."
        '
        'СправочникиToolStripMenuItem
        '
        Me.СправочникиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПолучателиToolStripMenuItem, Me.ТехникаToolStripMenuItem1})
        Me.СправочникиToolStripMenuItem.Name = "СправочникиToolStripMenuItem"
        Me.СправочникиToolStripMenuItem.Size = New System.Drawing.Size(94, 20)
        Me.СправочникиToolStripMenuItem.Text = "Справочники"
        '
        'ПолучателиToolStripMenuItem
        '
        Me.ПолучателиToolStripMenuItem.Name = "ПолучателиToolStripMenuItem"
        Me.ПолучателиToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.ПолучателиToolStripMenuItem.Text = "Получатели"
        '
        'ТехникаToolStripMenuItem1
        '
        Me.ТехникаToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыдатьToolStripMenuItem, Me.ДобавитьToolStripMenuItem, Me.УдалитьToolStripMenuItem})
        Me.ТехникаToolStripMenuItem1.Name = "ТехникаToolStripMenuItem1"
        Me.ТехникаToolStripMenuItem1.Size = New System.Drawing.Size(141, 22)
        Me.ТехникаToolStripMenuItem1.Text = "Техника"
        '
        'ВыдатьToolStripMenuItem
        '
        Me.ВыдатьToolStripMenuItem.Name = "ВыдатьToolStripMenuItem"
        Me.ВыдатьToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.ВыдатьToolStripMenuItem.Text = "Выдать..."
        '
        'ДобавитьToolStripMenuItem
        '
        Me.ДобавитьToolStripMenuItem.Name = "ДобавитьToolStripMenuItem"
        Me.ДобавитьToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.ДобавитьToolStripMenuItem.Text = "Добавить..."
        '
        'УдалитьToolStripMenuItem
        '
        Me.УдалитьToolStripMenuItem.Name = "УдалитьToolStripMenuItem"
        Me.УдалитьToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.УдалитьToolStripMenuItem.Text = "Удалить..."
        '
        'СправкаToolStripMenuItem
        '
        Me.СправкаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОПрограммеToolStripMenuItem, Me.ОАвтореToolStripMenuItem})
        Me.СправкаToolStripMenuItem.Name = "СправкаToolStripMenuItem"
        Me.СправкаToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.СправкаToolStripMenuItem.Text = "Справка"
        '
        'ОПрограммеToolStripMenuItem
        '
        Me.ОПрограммеToolStripMenuItem.Name = "ОПрограммеToolStripMenuItem"
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(158, 22)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе..."
        '
        'ОАвтореToolStripMenuItem
        '
        Me.ОАвтореToolStripMenuItem.Name = "ОАвтореToolStripMenuItem"
        Me.ОАвтореToolStripMenuItem.Size = New System.Drawing.Size(158, 22)
        Me.ОАвтореToolStripMenuItem.Text = "О авторе..."
        '
        'ВидToolStripMenuItem
        '
        Me.ВидToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.БольшиеИконкиToolStripMenuItem, Me.ДеталиToolStripMenuItem, Me.МаленькиеИконкиToolStripMenuItem, Me.СписокToolStripMenuItem, Me.НазваниеToolStripMenuItem})
        Me.ВидToolStripMenuItem.Name = "ВидToolStripMenuItem"
        Me.ВидToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.ВидToolStripMenuItem.Text = "Вид"
        '
        'БольшиеИконкиToolStripMenuItem
        '
        Me.БольшиеИконкиToolStripMenuItem.Name = "БольшиеИконкиToolStripMenuItem"
        Me.БольшиеИконкиToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.БольшиеИконкиToolStripMenuItem.Text = "Большие иконки"
        '
        'ДеталиToolStripMenuItem
        '
        Me.ДеталиToolStripMenuItem.Name = "ДеталиToolStripMenuItem"
        Me.ДеталиToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ДеталиToolStripMenuItem.Text = "Детали"
        '
        'МаленькиеИконкиToolStripMenuItem
        '
        Me.МаленькиеИконкиToolStripMenuItem.Name = "МаленькиеИконкиToolStripMenuItem"
        Me.МаленькиеИконкиToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.МаленькиеИконкиToolStripMenuItem.Text = "Маленькие иконки"
        '
        'СписокToolStripMenuItem
        '
        Me.СписокToolStripMenuItem.Name = "СписокToolStripMenuItem"
        Me.СписокToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.СписокToolStripMenuItem.Text = "Список"
        '
        'НазваниеToolStripMenuItem
        '
        Me.НазваниеToolStripMenuItem.Name = "НазваниеToolStripMenuItem"
        Me.НазваниеToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.НазваниеToolStripMenuItem.Text = "Название"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.AllowDrop = True
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыдатьToolStripMenuItem1, Me.ToolStripMenuItem2, Me.ЖурналВыдачДанногоУстройстваToolStripMenuItem, Me.ТекущийОтчетДанногоУстройстваToolStripMenuItem, Me.УдалитьToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(269, 98)
        '
        'ВыдатьToolStripMenuItem1
        '
        Me.ВыдатьToolStripMenuItem1.Name = "ВыдатьToolStripMenuItem1"
        Me.ВыдатьToolStripMenuItem1.Size = New System.Drawing.Size(268, 22)
        Me.ВыдатьToolStripMenuItem1.Text = "Выдать"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(265, 6)
        '
        'ЖурналВыдачДанногоУстройстваToolStripMenuItem
        '
        Me.ЖурналВыдачДанногоУстройстваToolStripMenuItem.Name = "ЖурналВыдачДанногоУстройстваToolStripMenuItem"
        Me.ЖурналВыдачДанногоУстройстваToolStripMenuItem.Size = New System.Drawing.Size(268, 22)
        Me.ЖурналВыдачДанногоУстройстваToolStripMenuItem.Text = "Журнал выдач данного устройства"
        '
        'ТекущийОтчетДанногоУстройстваToolStripMenuItem
        '
        Me.ТекущийОтчетДанногоУстройстваToolStripMenuItem.Name = "ТекущийОтчетДанногоУстройстваToolStripMenuItem"
        Me.ТекущийОтчетДанногоУстройстваToolStripMenuItem.Size = New System.Drawing.Size(268, 22)
        Me.ТекущийОтчетДанногоУстройстваToolStripMenuItem.Text = "Текущий отчет данного устройства"
        '
        'УдалитьToolStripMenuItem1
        '
        Me.УдалитьToolStripMenuItem1.Name = "УдалитьToolStripMenuItem1"
        Me.УдалитьToolStripMenuItem1.Size = New System.Drawing.Size(268, 22)
        Me.УдалитьToolStripMenuItem1.Text = "Удалить"
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВидToolStripMenuItem1, Me.ОбновитьToolStripMenuItem1, Me.ДобавитьToolStripMenuItem1})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(129, 70)
        '
        'ВидToolStripMenuItem1
        '
        Me.ВидToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.БольшиеИконкиToolStripMenuItem1, Me.ДеталиToolStripMenuItem1, Me.МаленькиеИконкиToolStripMenuItem1, Me.СписокToolStripMenuItem1, Me.НазваниеToolStripMenuItem1})
        Me.ВидToolStripMenuItem1.Name = "ВидToolStripMenuItem1"
        Me.ВидToolStripMenuItem1.Size = New System.Drawing.Size(128, 22)
        Me.ВидToolStripMenuItem1.Text = "Вид"
        '
        'БольшиеИконкиToolStripMenuItem1
        '
        Me.БольшиеИконкиToolStripMenuItem1.Name = "БольшиеИконкиToolStripMenuItem1"
        Me.БольшиеИконкиToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.БольшиеИконкиToolStripMenuItem1.Text = "Большие иконки"
        '
        'ДеталиToolStripMenuItem1
        '
        Me.ДеталиToolStripMenuItem1.Name = "ДеталиToolStripMenuItem1"
        Me.ДеталиToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.ДеталиToolStripMenuItem1.Text = "Детали"
        '
        'МаленькиеИконкиToolStripMenuItem1
        '
        Me.МаленькиеИконкиToolStripMenuItem1.Name = "МаленькиеИконкиToolStripMenuItem1"
        Me.МаленькиеИконкиToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.МаленькиеИконкиToolStripMenuItem1.Text = "Маленькие иконки"
        '
        'СписокToolStripMenuItem1
        '
        Me.СписокToolStripMenuItem1.Name = "СписокToolStripMenuItem1"
        Me.СписокToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.СписокToolStripMenuItem1.Text = "Список"
        '
        'НазваниеToolStripMenuItem1
        '
        Me.НазваниеToolStripMenuItem1.Name = "НазваниеToolStripMenuItem1"
        Me.НазваниеToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.НазваниеToolStripMenuItem1.Text = "Название"
        '
        'ОбновитьToolStripMenuItem1
        '
        Me.ОбновитьToolStripMenuItem1.Name = "ОбновитьToolStripMenuItem1"
        Me.ОбновитьToolStripMenuItem1.Size = New System.Drawing.Size(128, 22)
        Me.ОбновитьToolStripMenuItem1.Text = "Обновить"
        '
        'ДобавитьToolStripMenuItem1
        '
        Me.ДобавитьToolStripMenuItem1.Name = "ДобавитьToolStripMenuItem1"
        Me.ДобавитьToolStripMenuItem1.Size = New System.Drawing.Size(128, 22)
        Me.ДобавитьToolStripMenuItem1.Text = "Добавить"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(523, 339)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Инвентаризация"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListView1 As ListView
    Friend WithEvents ImageList2 As ImageList
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ФайлToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПерезагрузкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОбновитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ВыходToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВидToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents БольшиеИконкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДеталиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents МаленькиеИконкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СписокToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НазваниеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИнструментыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтчётПоВыданнойТехникиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЖурналToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправочникиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПолучателиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОПрограммеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОАвтореToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ВыдатьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents ЖурналВыдачДанногоУстройстваToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТекущийОтчетДанногоУстройстваToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УдалитьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents ВидToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ОбновитьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ДобавитьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents БольшиеИконкиToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ДеталиToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents МаленькиеИконкиToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents СписокToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents НазваниеToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ТехникаToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ВыдатьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДобавитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УдалитьToolStripMenuItem As ToolStripMenuItem
End Class
