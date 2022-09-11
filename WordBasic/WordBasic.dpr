program WordBasic;

uses
  ComObj,
  System.Variants,
  ActiveX,
  System.SysUtils;

var
   WB : OleVariant;
   I : integer;
   today : TDateTime;

begin
   today := Now;
   CoInitialize(nil);
   DeleteFile(PChar(GetCurrentDir + '\' + DateToStr(today) + '.doc'));
   try
      try
         WB := GetActiveOleObject('Word.Basic');
      except
         WB := CreateOleObject('Word.Basic');
      end;
      WB.FileNew('Normal');
      WB.AppShow;
      WB.Font('Times New Roman', 14);
      WB.Insert('Author: Kliment Lagrangiewicz'#13);

      WB.TableInsertTable(NumColumns := 8, NumRows := 6, Format := 7);
      WB.TableSelectTable;
      WB.Insert('First column');
      WB.NextCell;
      WB.Insert('Second column');
      WB.NextCell;
      WB.Insert('Third column');
      WB.NextCell;
      WB.Insert('Fourth column');
      WB.NextCell;
      WB.Insert('Fifth column');
      WB.NextCell;
      WB.Insert('Sixth column');
      WB.NextCell;
      WB.Insert('Seventh column');
      WB.NextCell;
      WB.Insert('Eighth column');
      WB.NextCell;
      for I := 1 to 40 do
      begin
         WB.Insert(IntToStr(I * I));
         if I <> 40 then
            WB.NextCell;
      end;

      WB.FileSaveAs(GetCurrentDir + '\' + DateToStr(today) + '.doc');
      WB.FileClose();
      WB.FileExit(1);
      if not VarIsEmpty(WB) then
         WB := Unassigned;
   except
      on E: Exception do
      begin
         Writeln(E.ClassName, ': ', E.Message);
         Sleep(2000);
      end;
   end;
   CoUninitialize();
   Sleep(1000);
end.
