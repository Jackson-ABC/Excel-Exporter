using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

public class VBA_Handling
{
    // TODO: Move fileType to a config file
    private string fileType = "vba";

    public void SetFileType(string fileType)
    {
        this.fileType = fileType;
    }

    public void ExportWorksheetVBA(Excel.Worksheet worksheet, string outputDir)
    {
        // Get the VBA project
        VBA.VBProject vbProject = worksheet.Parent.VBProject;
        if (vbProject != null)
        {
            // Get the VBA components
            VBA.VBComponents vbComponents = vbProject.VBComponents;
            
            // Create directory for VBA code
            string vbaDir = Path.Combine(outputDir, "Microsoft Excel Objects");
            Directory.CreateDirectory(vbaDir);

            // Get the worksheet's code module
            
            VBA.VBComponent vbComponent = vbComponents.Item(worksheet.CodeName);
            if (vbComponent != null)
            {
                // Create a .bas file for the worksheet's code
                string vbaFilePath = Path.Combine(vbaDir, $"{worksheet.CodeName} ({worksheet.Name}).{fileType}");

                // Write the VBA code to file
                int lineCount = vbComponent.CodeModule.CountOfLines;
                if (lineCount > 0)
                    File.WriteAllText(vbaFilePath, vbComponent.CodeModule.Lines[1, lineCount]);
            }
        }
    }

    public void ExportThisWorkbookVBA(Excel.Workbook workbook, string outputDir)
    {
        // Get the VBA project
        VBA.VBProject vbProject = workbook.VBProject;
        if (vbProject != null)
        {
            // Get the VBA components
            VBA.VBComponents vbComponents = vbProject.VBComponents;
            
            // Create directory for VBA code
            string vbaDir = Path.Combine(outputDir, "ThisWorkbook");
            Directory.CreateDirectory(vbaDir);

            // Get the ThisWorkbook code module
            VBA.VBComponent vbComponent = vbComponents.Item("ThisWorkbook");
            if (vbComponent != null)
            {
                // Create a .bas file for the ThisWorkbook code
                string vbaFilePath = Path.Combine(vbaDir, $"ThisWorkbook.{fileType}");

                // Write the VBA code to file
                int lineCount = vbComponent.CodeModule.CountOfLines;
                if (lineCount>0)
                    File.WriteAllText(vbaFilePath, vbComponent.CodeModule.Lines[1, lineCount]);
            }
        }
    }

    public void ExportModulesVBA(Excel.Workbook workbook, string outputDir)
    {
        // Get the VBA project
        VBA.VBProject vbProject = workbook.VBProject;
        if (vbProject != null)
        {
            // Get the VBA components
            VBA.VBComponents vbComponents = vbProject.VBComponents;
            
            // Create directory for VBA code
            string vbaDir = Path.Combine(outputDir, "Modules");
            Directory.CreateDirectory(vbaDir);

            // Get the modules
            foreach (VBA.VBComponent vbComponent in vbComponents)
            {
                if (vbComponent.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)
                {
                    // Create a file for the module
                    string vbaFilePath = Path.Combine(vbaDir, $"{vbComponent.Name}.{fileType}");

                    // Write the VBA code to file
                    int lineCount = vbComponent.CodeModule.CountOfLines;
                    if (lineCount > 0)
                        File.WriteAllText(vbaFilePath, vbComponent.CodeModule.Lines[1, lineCount]);
                }
            }
        }
    }

    public void ExportClassesVBA(Excel.Workbook workbook, string outputDir)
    {
        // Get the VBA project
        VBA.VBProject vbProject = workbook.VBProject;
        if (vbProject != null)
        {
            // Get the VBA components
            VBA.VBComponents vbComponents = vbProject.VBComponents;
            
            // Create directory for VBA code
            string vbaDir = Path.Combine(outputDir, "Classes");
            Directory.CreateDirectory(vbaDir);

            // Get the classes
            foreach (VBA.VBComponent vbComponent in vbComponents)
            {
                if (vbComponent.Type == VBA.vbext_ComponentType.vbext_ct_ClassModule)
                {
                    // Create a file for the class
                    string vbaFilePath = Path.Combine(vbaDir, $"{vbComponent.Name}.{fileType}");

                    // Write the VBA code to file
                    int lineCount = vbComponent.CodeModule.CountOfLines;
                    if (lineCount > 0)
                        File.WriteAllText(vbaFilePath, vbComponent.CodeModule.Lines[1, lineCount]);
                }
            }
        }
    }

    public void ExportFormsVBA(Excel.Workbook workbook, string outputDir)
    {
        // Get the VBA project
        VBA.VBProject vbProject = workbook.VBProject;
        if (vbProject != null)
        {
            // Get the VBA components
            VBA.VBComponents vbComponents = vbProject.VBComponents;
            
            // Create directory for VBA code
            string vbaDir = Path.Combine(outputDir, "Forms");
            Directory.CreateDirectory(vbaDir);

            // Get the forms
            foreach (VBA.VBComponent vbComponent in vbComponents)
            {
                // TODO: Correctly implement using the below vba code for exporting frx files
                // Currently generates a directory for the form but does not export the frx file
                if (vbComponent.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                {
                    // Create a directory for the form
                    Directory.CreateDirectory(Path.Combine(vbaDir, vbComponent.Name));
                    
                    // Create a file for the form
                    // string vbaFilePath = Path.Combine(Path.Combine(vbaDir, vbComponent.Name), $"{vbComponent.Name}.{fileType}");
                    
                    // Write the VBA code to file
                    // File.WriteAllText(vbaFilePath, vbComponent.CodeModule.Lines[1, vbComponent.CodeModule.CountOfLines]);
                }
            }
        }
    }
}

/*
    ' vba code to extract form files
    #COMPILE EXE

    GLOBAL gsDumpDir AS STRING


    SUB SaveFile (sFilename AS STRING, sData AS STRING)
    LOCAL hFile AS DWORD
    hFile = FREEFILE
    OPEN gsDumpDir & sFilename FOR BINARY ACCESS WRITE AS #hFile
    PUT #hFile, 1, sData
    CLOSE #hFile
    END SUB


    FUNCTION FRXGetText(sFRXFile AS STRING, BYVAL dwOffset AS DWORD) AS STRING
    LOCAL hFile AS DWORD, sBuf AS STRING, bPtr AS BYTE PTR, wPtr AS WORD PTR, dwPtr AS DWORD PTR
    hFile = FREEFILE
    OPEN sFRXFile FOR BINARY ACCESS READ LOCK SHARED AS #hFile BASE = 0
    SEEK #hFile, dwOffset
    GET$ #hFile, LOF(hFile), sBuf
    CLOSE #hFile
    bPtr = STRPTR(sBuf)
    IF @bPtr = &hFF THEN
        wPtr = bPtr + 1
        FUNCTION = MID$(sBuf, 4, @wPtr)
    ELSEIF @bPtr = &h00 THEN
        dwPtr = bPtr
        FUNCTION = MID$(sBuf, 5, @dwPtr)
    ELSE
        FUNCTION = MID$(sBuf, 2, @bPtr)
    END IF
    END FUNCTION


    FUNCTION FRXGetList(sFRXFile AS STRING, BYVAL dwOffset AS DWORD) AS STRING
    LOCAL hFile AS DWORD, sBuf AS STRING, wPtr AS WORD PTR, dwListItems AS DWORD, i AS LONG, sList AS STRING
    hFile = FREEFILE
    OPEN sFRXFile FOR BINARY ACCESS READ LOCK SHARED AS #hFile BASE = 0
    SEEK #hFile, dwOffset
    GET$ #hFile, LOF(hFile), sBuf
    CLOSE #hFile
    wPtr = STRPTR(sBuf)
    dwListItems = @wPtr
    wPtr = wPtr + 4
    FOR i = 1 TO dwListItems
        sList = sList & MID$(sBuf, wPtr - STRPTR(sBuf) + 3, @wPtr) & $CRLF
        wPtr = wPtr + @wPtr + 2
    NEXT i
    FUNCTION = sList
    END FUNCTION


    FUNCTION FRXGetBinary(sFRXFile AS STRING, BYVAL dwOffset AS DWORD) AS STRING
    LOCAL hFile AS DWORD, sBuf AS STRING, dwPtr AS DWORD PTR, sData AS STRING
    hFile = FREEFILE
    OPEN sFRXFile FOR BINARY ACCESS READ LOCK SHARED AS #hFile BASE = 0
    SEEK #hFile, dwOffset
    GET$ #hFile, LOF(hFile), sBuf
    CLOSE #hFile
    dwPtr = STRPTR(sBuf) + 8
    sData = MID$(sBuf, 13, @dwPtr)
    FUNCTION = sData
    END FUNCTION


    FUNCTION GetFileExt(sData AS STRING) AS STRING
    SELECT CASE LEFT$(sData,2)
    CASE CHR$(&h42,&h4D): FUNCTION = ".bmp"
    CASE CHR$(&hFF,&hD8): FUNCTION = ".jpg"
    CASE CHR$(&h47,&h49): FUNCTION = ".gif"
    CASE CHR$(&h49,&h49): FUNCTION = ".tif"
    CASE CHR$(&h89,&h50): FUNCTION = ".png"
    CASE CHR$(&hD7,&hCD): FUNCTION = ".wmf"
    CASE ELSE: FUNCTION = ".dat"
    END SELECT
    END FUNCTION


    SUB DumpVBForm (sVBFile AS STRING)
    LOCAL sFRMFile AS STRING, sFRXFile AS STRING, hFile AS DWORD, sLine AS STRING, i AS LONG, sOffset AS STRING
    LOCAL sCurObj AS STRING, sCurVal AS STRING, sData AS STRING, sObject AS STRING
    sFRMFile = LEFT$(sVBFile, LEN(sVBFile) - 1) & "m"
    sFRXFile = LEFT$(sVBFile, LEN(sVBFile) - 1) & "x"
    MID$(sFRMFile, LEN(sFRMFile) - 2, 3) = LCASE$(RIGHT$(sFRMFile,3))
    MID$(sFRXFile, LEN(sFRXFile) - 2, 3) = LCASE$(RIGHT$(sFRXFile,3))
    hFile = FREEFILE
    ERRCLEAR
    OPEN sFRMFile FOR INPUT LOCK SHARED AS #hFile
    IF ERR THEN
        ? "ERROR: Couldn't open " & sFRMFile: EXIT SUB
    END IF
    LINE INPUT #hFile, sLine
    IF sLine <> "VERSION 5.00" THEN    '// 5.00 = VB5 & VB6
        ? "ERROR: FRM file version invalid (not VB5/6)": EXIT SUB
    END IF
    DO UNTIL EOF(hFile)
        LINE INPUT #hFile, sLine
        sLine = TRIM$(sLine)
        IF sLine = "" THEN ITERATE
        DO
            IF INSTR(1, sLine, "  ") = 0 THEN EXIT DO
            REPLACE "  " WITH " " IN sLine
        LOOP
        IF LEFT$(sLine,6) = "Begin " THEN
            sObject = sObject & RIGHT$(sLine, LEN(sLine) - INSTR(-1, sLine, " ")) & "."
        ELSEIF sLine = "End" THEN
            sObject = LEFT$(sObject, INSTR(-2, sObject, "."))
        ELSEIF sLine = "Attribute" THEN 'Ignored
        ELSE
            IF sObject = "" THEN ITERATE
            sCurObj = LEFT$(sLine, INSTR(1, sLine, " =") - 1)
            sCurVal = RIGHT$(sLine, LEN(sLine) - INSTR(1, sLine, " =") - 2)
            IF RIGHT$(sCurVal,1) <> CHR$(34) THEN
            IF INSTR(1, sCurVal, ":") THEN
            sOffset = RIGHT$(sCurVal, LEN(sCurVal) - INSTR(-1, sCurVal, ":"))
            IF sOffset <> "" THEN
                SELECT CASE sCurObj
                CASE "Text": SaveFile sObject & sCurObj & ".txt", FRXGetText(sFRXFile, BYVAL VAL("&h0" & sOffset))
                CASE "List", "ItemData": SaveFile sObject & sCurObj & ".txt", FRXGetList(sFRXFile, BYVAL VAL("&h0" & sOffset))
                CASE "Picture", "DisabledPicture", "DownPicture", "Palette": sData = FRXGetBinary(sFRXFile, BYVAL VAL("&h0" & sOffset))
                                        SaveFile sObject & sCurObj & GetFileExt(sData), sData
                CASE "Icon", "MouseIcon", "DragIcon": sData = FRXGetBinary(sFRXFile, BYVAL VAL("&h0" & sOffset))
                                        SaveFile sObject & sCurObj & ".ico", sData
                CASE ELSE: sData = FRXGetBinary(sFRXFile, BYVAL VAL("&h0" & sOffset))
                                        SaveFile sObject & sCurObj & ".dat", sData
                END SELECT
            END IF
            END IF
            ELSE
                IF sCurObj = "Text" THEN SaveFile sObject & "Text.txt", MID$(sCurVal, 2, LEN(sCurVal) - 2)
            END IF
        END IF
    LOOP
    CLOSE #hFile
    END SUB


    FUNCTION PBMAIN () AS LONG
    LOCAL sVBForm AS STRING, i AS LONG

    '// Specify the VB Form to extract files from
    sVBForm = "e:\dev\vb98\temp\frx\multi\form1.frx"   '// .FRM or .FRX
    
    '// Specify directory to extract to, and ensure it exists
    gsDumpDir = "c:\frxdump"
    IF RIGHT$(gsDumpDir,1) = "\" THEN gsDumpDir = LEFT$(gsDumpDir, LEN(gsDumpDir) - 1)
    MKDIR gsDumpDir
    gsDumpDir = gsDumpDir & "\"

    '// Extract files from the VB Form
    DumpVBForm sVBForm

    '// Show the results
    i = SHELL("explorer " & gsDumpDir)
    END FUNCTION
*/