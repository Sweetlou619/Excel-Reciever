   M S B A M L   `   `   `  ����  Excel Retriever     Excel_Retriever.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   a ]Telerik.Windows.Data, Version=2012.1.326.40, Culture=neutral, PublicKeyToken=5803cfa389c90ce7n jTelerik.Windows.Controls.Charting, Version=2012.1.326.40, Culture=neutral, PublicKeyToken=5803cfa389c90ce7e aTelerik.Windows.Controls, Version=2012.1.326.40, Culture=neutral, PublicKeyToken=5803cfa389c90ce7k gTelerik.Windows.Controls.Chart, Version=2012.1.326.40, Culture=neutral, PublicKeyToken=5803cfa389c90ce7Etelerik1http://schemas.telerik.com/2008/xaml/presentation     5   	     � Title$   ��5   	   $	��591��6   $	��903��6   .��5      � -   5        $��	ExcelGrid��5      $y�Visible=�6   .+�5   
   �� $y�	Collapsed=�6   .+�5	      '  !Telerik.Windows.Controls.RadChart  -   5          Name$ graph��5	      $��1��6.   $��1��68     5
      -   5        $ BubbleGraph��5
      $��1��64   $��1��6>   5      �� 5   
   -   5        $��
CombineBtn��5   \   $"��Combine Comments and Ratings��6   {�.  $��12,433,682,74q�6�   �� 5   
   -   5        $��	ExportBtn��5   K   $��Export Bar Graphs��6   {�.  $��12,375,682,134q�6n   �� 5   
   -   5        $��ExportBubbleBtn��5   R   $��Export Bubble Charts��6   {�.  $��12,497,682,12q�6{   �� 5   
   -   5        $��BeginSheetBtn��5   Q   $"��Import Beginning Excel Sheet��6   $��12,145,682,361q�6f   �� 5   
   -   5        $��EndSheetBtn��5   I   $��Import End Excel Sheet��6   $��12,313,682,192q�6\   �� 5   
   -	   5        $��LeftToExportText��5       $��23��6   $��213,395,0,0q�68   $	��Top=�6M   $
��Left=�6e   $	��182��6�   �� 5   
   -
   5        $��CombineText��5       $��23��6   $��213,455,486,74q�63   �� 5   
   -   5        $��ExportBubbleText��5       $��23��6   $��213,517,486,12q�68   �� 5   
   -   5        $��BeginImportText��5       $��23��6   $��213,168,486,361q�67   �� 5   
   -   5        $��EndImportText��5       $��23��6   $��213,337,486,192q�65   �� 5   
   -   5        $��ColumnBlockBegin��5   N   $��23��6   $
��Left=�6    $��	12,44,0,0q�6;   $,��&How many columns with names are there?��6f   $	��Top=�6�   �� 5   
   -   5        $��CompletelyBlockBegin��5   N   $��23��6   $
��Left=�6    $��	12,76,0,0q�6;   $"��What row is 'Completely' in?��6j   $	��Top=�6�   �� 5   
   -   5        $��ColumnBlockEnd��5   O   $��23��6   $
��Left=�6    $��
12,207,0,0q�6;   $,��&How many columns with names are there?��6e   $	��Top=�6�   �� 5   
   -   5        $��CompletelyBlockEnd��5   O   $��23��6   $
��Left=�6    $��
12,247,0,0q�6;   $"��What row is 'Completely' in?��6i   $	��Top=�6�   �� 5   
   -   5        $��NumNamesRowBlockEnd��5   O   $��23��6   $
��Left=�6    $��
12,284,0,0q�6;   $#��What row is the last name in?��6j   $	��Top=�6�   �� 5   
   -   5        $��NumNamesRowBlockBegin��5   O   $��23��6   $
��Left=�6    $��
12,105,0,0q�6;   $#��What row is the last name in?��6l   $	��Top=�6�   �� 5   
   -   5        $��NumColumnsBegin��5   M   $��23��6   $
��Left=�6   $��
243,41,0,0q�69   $	��Top=�6d   $��45��6|   $��1��6�   �� 5   
   -   5        $��CompletelyStartBegin��5   M   $��23��6   $
��Left=�6   $��
168,73,0,0q�69   $	��Top=�6i   $��45��6�   $��1��6�   �� 5   
   -   5        $��NumNamesRowBeginBox��5   N   $��23��6   $
��Left=�6   $��174,102,0,0q�69   $	��Top=�6i   $��45��6�   $��1��6�   �� 5   
   -   5        $��NumColumnsEnd��5   N   $��23��6   $
��Left=�6   $��243,204,0,0q�69   $	��Top=�6c   $��45��6{   $��1��6�   �� 5    
   -   5        $��CompletelyStartEnd��5    N   $��23��6   $
��Left=�6   $��168,244,0,0q�69   $	��Top=�6h   $��45��6�   $��1��6�   �� 5!   
   -   5        $��NumNamesRowEndBox��5!   N   $��23��6   $
��Left=�6   $��174,281,0,0q�69   $	��Top=�6g   $��45��6   $��1��6�   �� 5"   
   -   5        $��ImportTwoGraphs��5"   l   $��Importing Two Graphs��6   $��28��62   $
��Left=�6>   $��	12,12,0,0q�6Y   $	��Top=�6�   $	��152��6�   6�   �� 5#   
   -   5        $��DisableNameBox��5#   s   $ ��Disable 'Name' message box��6   $��28��68   $
��Left=�6D   $��
186,12,0,0q�6_   $	��Top=�6�   $	��181��6�   6�   5C      5D      