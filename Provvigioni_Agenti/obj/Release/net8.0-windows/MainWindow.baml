   M S B A M L   `   `   `  ���� M  IProvvigioni_Agenti, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:Provvigioni_AgentiProvvigioni_Agenti  #    Provvigioni_Agenti.MainWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     2local clr-namespace:Provvigioni_Agenti       � Title$  
MainWindow��$	��800��$
��1290��.���  J� RenderTransformOrigin$ 1,0.5&�$��0,0,4,0q�z�V� .~��� �� �� T�  T� X$ 2Q� � ColumnDefinitions �� $	��120 ��� $	��120 ��� $	��150 ��� $	��580 ��� $	��300 � � RowDefinitions �� $��40 ��� $	��140 ��� $��2* ��� $��7* ��� $	��10* �.+��� $��0�� �� Orientation$ 
Horizontal=�$��Center=�$��Center=�$��30��$	��970��$��4��.+��� -    ��Name$ 
agenteInfo��$��15(�$��-��$��	60,10,0,0q�$
��Bold"��� $��1��$��2��$ Vertical=�$
��Left=�$	��Top=�$��2��$	��150��$	��221��.+��� $ 
Horizontal=�$��	10,10,0,0q�.+��� $��agente��$��Center=�$��0,0,10,0q�$��40��$��Center=��� -    ��Name$ elencoAgenti��$	��150���� $ 
Horizontal=�$��	10,10,0,0q�.+��� $	��dal��$��0,0,10,0q�$��40��$��Right=�$��Center=�(  "System.Windows.Controls.DatePicker  -     Name$ fromDate��$	��150���� $ 
Horizontal=�$��	10,10,0,0q�.+��� $��al��$��Right=�$��0,0,10,0q�$��40��$��Center=�  -   $ toDate��$	��150���� $��10q�.+��� -   	 ��Name$	 buttonElabora��$��30��$	��201��.��	ELABORA�� $��1��$��2��$ Vertical=�$��	10,10,0,0q�.+�	� -   
 	�Name$	
 t_1��$��2q�.��1° TRIMESTRE	� -   $	
 t_2��$��2q�.��2° TRIMESTRE	� -   $	
 t_3��$��2q�.��3° TRIMESTRE	� -	   $	
 t_4��$��2q�.��4° TRIMESTRE�� $ 
Horizontal=�$��0,5,0,0q�.+��� $	��DEL��$��0,0,10,0q�$��30���� -
   $ annoCorrente��$��80���� $ 
Horizontal=�$��0,5,0,0q�.+��� $
��PREC��$��0,0,10,0q�$��30���� -   $ annoRiferimento��$��80���� $��5��$��1��$��1��$ Vertical=�$��0,0,0,0q�$
��Left=�$��Bottom=�$	��125��$	��250��.+��� $ Vertical=�.+��� $��20(�$��PROVVIGIONI TOTALI��$��0,0,10,0q�$
��Bold"� �� TextAlignment$ Center=��� -   $ totProvvigioneTrimestre��$��20(�$	��---��$��0,5,10,0q�$
��Bold"�$ Center=��� $��Bottom=�$��44q�.+��� -   $��apri directory��$��82��$��20���� $��3��$��1��$��2��$ Vertical=�$��0,0,0,0q�$
��Left=�$��Bottom=�$	��125��$	��575��.+��� $ 
Horizontal=�.+��� $ 
Horizontal=�$��0,5,0,0q�.+��� $��0,0,10,0q�$��15(�$��50���� -   $ titoloTrimestre��$��15(�$��	Trimestre��$��0,0,10,0q�$
��Bold"�$	��110��$ Center=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ titoloAnnuale��$��15(�$��Progressivo��$��0,0,10,0q�$
��Bold"�$	��110��$ Center=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ provvigioneTrimestre��$��15(�$��Provvigione��$��0,0,10,0q�$
��Bold"�$	��110��$ Center=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $! provvigioneTrimestreSellout��$��15(�$��Provv. Sellout��$��0,0,10,0q�$
��Bold"�$	��110��$ Center=��� $ 
Horizontal=�.+��� $ Vertical=�.+��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $
 corr��$��0,0,10,0q�$��15(�$��50���� -   $ totInfo��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $	 rif��$��0,0,10,0q�$��15(�$��50���� -   $ totInfoPrec��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� $��Δ��$��0,0,10,0q�$��15(�$��50���� -   $ deltaTrimestre��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� $
��Δ %��$��0,0,10,0q�$��15(�$��50���� -   $ deltaTrimestrePercent��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ Vertical=�.+��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $  totInfoProgressivoCorrente��$��15(�$	��---��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $# totInfoProgressivoRiferimento��$��15(�$	��---��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ deltaProgressivo��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ deltaProgressivoPercent��$��15(�$�� ��$��0,0,10,0q�$
��Bold"�$	��110��$ Right=��� $ Vertical=�.+��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ totProvvigioneCorrente��$��15(�$	��---��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ daUtilizzare0��$��15(�$�� ��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=��� $ Vertical=�.+��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ totProvvigioneSellout��$��15(�$	��---��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=��� $ 
Horizontal=�$��0,5,0,0q�.+��� -   $ daUtilizzare��$��15(�$�� ��$	��110��$��0,0,10,0q�$
��Bold"�$ Right=�&   System.Windows.Controls.DataGrid  -      Name$ dataGridTrasferiti��$��5��$��5,5,5,10q�$��2��$��2��$
��True��$��13(�   	RowHeight$ 30Q��� $��5��$��4��$��2��$��5,5,5,10q�$ Vertical=�.+��� -!   $ nomeClienteCategoriaLabel��$��15(�$	��---��$
 Left=�$
��Bold"�  -"   $" dataGridCategorieStatCliente��$
��True��$��13(�$ 30Q�  -#   $ dataGridVendite��$��0��$��	10,5,5,10q�$��2��$��4��$��3��$
��True��$��13(�$ 30Q�   RowStyle ��  �� 
TargetType)  #System.Windows.Controls.DataGridRow"   �� Triggers P�  P� Property �@J�IsMouseOver P� Value$
 True��. ���  �� Property � �� �� Value
 ��  ���  � ��
 ������� -$   $	 creaExcelButton��$��3��$��
CREA EXCEL��$
��Left=�$	��160��$��319,110,0,0q�$ 0.379,-2.992&�$��1���� -%   $	 directoryExcelFinale��$��3��$��	directory��$
��Left=�$��86��$��20��$��488,120,0,0q�$	��Top=�$ 0.379,-2.992&�$��1���� $��5��$��SELL OUT��$
��Left=�$��	0,114,0,0q�$��1��$	��Top=�$
��Bold"�$ 2.163,-0.502&�