package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"os/exec"
	//"errors"
	"path"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable *RefTable
	Date1904       bool
	styles         *xlsxStyleSheet
	Sheets         []*Sheet
	Sheet          map[string]*Sheet
	theme          *theme
	DefinedNames   []*xlsxDefinedName
}

func (f *File) AddCF(cf map[string][]map[string]string) (err error){
	// loop over styles and get largest dxfid
	// after setting new dxfs increment count of dxfs
	newDxfId := "0"	
	if f.styles != nil {
		newDxfId = f.styles.Dxfs.Count
		if newDxfId == ""{
			newDxfId = "0"
		}
	}
	
	for _, sheet := range f.Sheets {		
		// loop over file CF to get dxf last priority
		lastPriority := 0
		for _, CFFile := range sheet.ConditionalFormatting {
			priorityCFFile, err := strconv.Atoi(CFFile.CfRule.Priority)
			if err != nil {
				return err
			}
			if priorityCFFile > lastPriority{
				lastPriority = priorityCFFile
			}
		}
		
		// loop over request cf's and add it to first sheet
		for _, CFMap := range cf["cf"]{			
			newCF := new(conditionalFormatting)
			newCF.Sqref = CFMap["sqref"]
			newCF.CfRule.Type = "expression"			
			newCF.CfRule.DxfId = newDxfId
			lastPriority++
			strLastPriority := strconv.Itoa(lastPriority)
			newCF.CfRule.Priority = strLastPriority
			newCF.CfRule.Formula.Value = CFMap["formula"]

			// add newCF to sheet
			sheet.ConditionalFormatting = append(sheet.ConditionalFormatting, *newCF)
			// make new xlsxFill from request
			dxfFill := new(xlsxFill)
			newDxf := new(xlsxDxf)	

			dxfFill.PatternFill.BgColor.RGB = CFMap["BgColor"]
			
			newDxf.Fill = *dxfFill
			if f.styles == nil{
				f.styles = new(xlsxStyleSheet)
			}
			f.styles.Dxfs.Dxf = append(f.styles.Dxfs.Dxf, *newDxf)
			
			intNewDxfId, err := strconv.Atoi(newDxfId)
			if err != nil {
				return err
			}
			
			intNewDxfId++
			newDxfId = strconv.Itoa(intNewDxfId)			
		}
		f.styles.Dxfs.Count = newDxfId
		break
	}
	return nil
}

// Create a new File
func NewFile() *File {
	return &File{
		Sheet:        make(map[string]*Sheet),
		Sheets:       make([]*Sheet, 0),
		DefinedNames: make([]*xlsxDefinedName, 0),
	}
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (file *File, err error) {
	var f *zip.ReadCloser
	f, err = zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	file, err = ReadZip(f)
	return
}

// OpenBinary() take bytes of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenBinary(bs []byte) (*File, error) {
	r := bytes.NewReader(bs)
	return OpenReaderAt(r, int64(r.Len()))
}

// OpenReaderAt() take io.ReaderAt of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenReaderAt(r io.ReaderAt, size int64) (*File, error) {
	file, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}
	return ReadZipReader(file)
}

// A convenient wrapper around File.ToSlice, FileToSlice will
// return the raw data contained in an Excel XLSX file as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func FileToSlice(path string) ([][][]string, error) {
	f, err := OpenFile(path)
	if err != nil {
		return nil, err
	}
	return f.ToSlice()
}

// Save the File to an xlsx file at the provided path.
func (f *File) Save(filePath string) (err error) {
	target, err := os.Create(filePath)
	if err != nil {
		return err
	}
	err = f.Write(target)
	if err != nil {
		return err
	}
	tCl := target.Close()
	for _, sheet := range f.Sheets {
		if len(sheet.ConditionalFormatting) != 0 {
			// repair xlsx with libreoffice
			// If worksheet has conditional formats repair the file with libreoffice for Excel 2011 on Mac  
		    _, err = exec.LookPath("libreoffice")
		    if err != nil {
		        strErr := "libreoffice utility is not installed! apt-get install libreoffice"
		        fmt.Println(strErr)
		        return err
		    }
		    fileName := path.Base(filePath)
		    path_without_filename := path.Dir(filePath)
		    if path_without_filename == "."{
		        path_without_filename = ""
		    }
			outdir := path_without_filename+"/repaired_xlsx"
		    if _, err := os.Stat(outdir); err != nil {
	        	if os.IsNotExist(err) {
		            if os.MkdirAll(outdir,0777) != nil {
		                return fmt.Errorf("Create "+outdir+" failed")
		            }
		        }
		    }
		    strCMD := `libreoffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" `+ filePath + " --outdir " + outdir 
		    out, err := exec.Command("/bin/sh", "-c", strCMD).Output()		    
		    fmt.Printf("Libreoffice xlsx repair for : %v ,out: %v \n", filePath, string(out))
		    if err != nil{
		        fmt.Printf("Libreoffice cmd err: %v", err.Error())
		        return err
		    }
		    overwritingSource := outdir+"/"+fileName
		    // cmd to overwrite input xlsx with output xlsx		   
		    strCMD = "mv -f "+overwritingSource+" "+filePath
		    out, err = exec.Command("/bin/sh", "-c", strCMD).Output()
		    if err != nil{
		    	fmt.Printf("Overwriting original xlsx failed : %v , out: %v \n", filePath, string(out))
		        fmt.Printf("Cmd err: %v", err.Error())
		        return err
		    }
			/*
		    //  remove repaired xlsx from output directory		   
		    strCMD = "rm "+overwritingSource
		    out, err = exec.Command("/bin/sh", "-c", strCMD).Output()	
		    fmt.Printf("Removing original xlsx failed for : %v , out: %v \n", filePath, string(out))
		    if err != nil{
		        fmt.Printf("Cmd err: %v", err.Error())
		        return err
		    }*/
		}
	}
	return tCl
}

// Write the File to io.Writer as xlsx
func (f *File) Write(writer io.Writer) (err error) {
	parts, err := f.MarshallParts()
	if err != nil {
		return
	}
	zipWriter := zip.NewWriter(writer)
	for partName, part := range parts {
		w, err := zipWriter.Create(partName)
		if err != nil {
			return err
		}
		_, err = w.Write([]byte(part))
		if err != nil {
			return err
		}
	}
	return zipWriter.Close()
}

// Add a new Sheet, with the provided name, to a File
func (f *File) AddSheet(sheetName string) (*Sheet, error) {
	if _, exists := f.Sheet[sheetName]; exists {
		return nil, fmt.Errorf("duplicate sheet name '%s'.", sheetName)
	}
	sheet := &Sheet{
		Name:     sheetName,
		File:     f,
		Selected: len(f.Sheets) == 0,
	}
	f.Sheet[sheetName] = sheet
	f.Sheets = append(f.Sheets, sheet)
	return sheet, nil
}

func (f *File) makeWorkbook() xlsxWorkbook {
	return xlsxWorkbook{
		FileVersion: xlsxFileVersion{AppName: "Go XLSX"},
		WorkbookPr:  xlsxWorkbookPr{ShowObjects: "all"},
		BookViews: xlsxBookViews{
			WorkBookView: []xlsxWorkBookView{
				{
					ShowHorizontalScroll: true,
					ShowSheetTabs:        true,
					ShowVerticalScroll:   true,
					TabRatio:             204,
					WindowHeight:         8192,
					WindowWidth:          16384,
					XWindow:              "0",
					YWindow:              "0",
				},
			},
		},
		Sheets: xlsxSheets{Sheet: make([]xlsxSheet, len(f.Sheets))},
		CalcPr: xlsxCalcPr{
			IterateCount: 100,
			RefMode:      "A1",
			Iterate:      false,
			IterateDelta: 0.001,
		},
	}

}

// Some tools that read XLSX files have very strict requirements about
// the structure of the input XML.  In particular both Numbers on the Mac
// and SAS dislike inline XML namespace declarations, or namespace
// prefixes that don't match the ones that Excel itself uses.  This is a
// problem because the Go XML library doesn't multiple namespace
// declarations in a single element of a document.  This function is a
// horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	newWorkbook := strings.Replace(workbookMarshal, `xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id`, `r:id`, -1)
	// Dirty hack to fix issues #63 and #91; encoding/xml currently
	// "doesn't allow for additional namespaces to be defined in the
	// root element of the document," as described by @tealeg in the
	// comments for #63.
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(newWorkbook, oldXmlns, newXmlns, 1)
}

// Construct a map of file name to XML content representing the file
// in terms of the structure of an XLSX file.
func (f *File) MarshallParts() (map[string]string, error) {
	var parts map[string]string
	var refTable *RefTable = NewSharedStringRefTable()
	refTable.isWrite = true
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.Marshal(thing)
		if err != nil {
			return "", err
		}
		return xml.Header + string(body), nil
	}

	parts = make(map[string]string)
	workbook = f.makeWorkbook()

	sheetIndex := 1

	if f.styles == nil {
		f.styles = newXlsxStyleSheet(f.theme)
	}
	f.styles.reset()
	for _, sheet := range f.Sheets {
		xSheet := sheet.makeXLSXSheet(refTable, f.styles)
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex-1] = xlsxSheet{
			Name:    sheet.Name,
			SheetId: sheetId,
			Id:      rId,
			State:   "visible"}
		parts[partName], err = marshal(xSheet)
		if err != nil {
			return parts, err
		}
		sheetIndex++
	}

	workbookMarshal, err := marshal(workbook)
	if err != nil {
		return parts, err
	}
	workbookMarshal = replaceRelationshipsNameSpace(workbookMarshal)
	parts["xl/workbook.xml"] = workbookMarshal
	if err != nil {
		return parts, err
	}

	parts["_rels/.rels"] = TEMPLATE__RELS_DOT_RELS
	parts["docProps/app.xml"] = TEMPLATE_DOCPROPS_APP
	// TODO - do this properly, modification and revision information
	parts["docProps/core.xml"] = TEMPLATE_DOCPROPS_CORE
	parts["xl/theme/theme1.xml"] = TEMPLATE_XL_THEME_THEME

	xSST := refTable.makeXLSXSST()
	parts["xl/sharedStrings.xml"], err = marshal(xSST)
	if err != nil {
		return parts, err
	}

	xWRel := workbookRels.MakeXLSXWorkbookRels()

	parts["xl/_rels/workbook.xml.rels"], err = marshal(xWRel)
	if err != nil {
		return parts, err
	}

	parts["[Content_Types].xml"], err = marshal(types)
	if err != nil {
		return parts, err
	}
	parts["xl/styles.xml"], err = f.styles.Marshal()
	if err != nil {
		return parts, err
	}

	return parts, nil
}

// Return the raw data contained in the File as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func (file *File) ToSlice() (output [][][]string, err error) {
	output = [][][]string{}
	for _, sheet := range file.Sheets {
		s := [][]string{}
		for _, row := range sheet.Rows {
			if row == nil {
				continue
			}
			r := []string{}
			for _, cell := range row.Cells {
				str, err := cell.String()
				if err != nil {
					return output, err
				}
				r = append(r, str)
			}
			s = append(s, r)
		}
		output = append(output, s)
	}
	return output, nil
}
