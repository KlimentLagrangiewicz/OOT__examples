program WordApplication;

uses
   ComObj,
   System.Variants,
   ActiveX,
   System.SysUtils;

procedure main;
var
   WA : OleVariant;
   I, J, K : integer;
   today : TDateTime;
begin
   today := Now;
   try
      DeleteFile(PChar(GetCurrentDir + '\' + DateToStr(today) + '.doc'));
      WA := CreateOleObject('Word.Application');
      WA.Visible := false;
      WA.Documents.Add();
      WA.Selection.Font.Name := 'Times New Roman';
      WA.Selection.Font.Size := 14;
      WA.Selection.TypeText('Author: Kliment Lagrangiewicz');

      WA.ActiveDocument.Tables.Add(WA.Selection.Range, 8, 8, EmptyParam, EmptyParam);
      WA.ActiveDocument.Tables.Item(1).Style := 'Таблица простая 4';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 1).Range.Text := 'First row and column';
      WA.ActiveDocument.Tables.Item(1).Cell(2, 1).Range.Text := 'Second row';
      WA.ActiveDocument.Tables.Item(1).Cell(3, 1).Range.Text := 'Third row';
      WA.ActiveDocument.Tables.Item(1).Cell(4, 1).Range.Text := 'Fourth row';
      WA.ActiveDocument.Tables.Item(1).Cell(5, 1).Range.Text := 'Fifth row';
      WA.ActiveDocument.Tables.Item(1).Cell(6, 1).Range.Text := 'Sixth row';
      WA.ActiveDocument.Tables.Item(1).Cell(7, 1).Range.Text := 'Seventh row';
      WA.ActiveDocument.Tables.Item(1).Cell(8, 1).Range.Text := 'Eighth row';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 2).Range.Text := 'Second column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 3).Range.Text := 'Third column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 4).Range.Text := 'Fourth column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 5).Range.Text := 'Fifth column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 6).Range.Text := 'Sixth column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 7).Range.Text := 'Seventh column';
      WA.ActiveDocument.Tables.Item(1).Cell(1, 8).Range.Text := 'Eighth column';
      K := 1;
      for I := 2 to 8 do
      begin
         for J := 2 to 8 do
         begin
            if ((I mod 2) = 0) then
               WA.ActiveDocument.Tables.Item(1).Cell(I, J).Range.Text := IntToStr(K * K)
            else
               WA.ActiveDocument.Tables.Item(1).Cell(I, 10 - J).Range.Text := IntToStr(K * K);
            inc(K);
         end;
      end;

      WA.ActiveDocument.SaveAs(GetCurrentDir + '\' + DateToStr(today) + '.doc');
      WA.ActiveDocument.Close();
      WA.Quit();
      WA := Unassigned;
   except
      on E: Exception do
      begin
         Writeln(E.ClassName, ': ', E.Message);
         Sleep(2000);
      end;
   end;
end;

begin
   OleCheck(CoInitialize(nil));
   main();
   CoUninitialize();
end.
