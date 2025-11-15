CREATE TABLE [TBLMENU] (
  [MenuGroupeNum] BYTE,
  [MenuNum] BYTE,
  [MenuLibelle] VARCHAR (50),
  [MenuMacroToRun] VARCHAR (100),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([MenuGroupeNum], [MenuNum])
)
