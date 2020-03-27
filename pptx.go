package pptx

import (
	"archive/zip"
	"fmt"
	"image/png"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/gen2brain/go-fitz"
)

// ConvertToPptx converts src to pptx and saves in dest
func ConvertToPptx(src, dest string) error {
	src, err := filepath.Abs(src)
	if err != nil {
		return err
	}
	dest, err = filepath.Abs(dest)
	if err != nil {
		return err
	}
	tempDir, err := ioutil.TempDir(os.TempDir(), "pdf2pptx")
	if err != nil {
		return err
	}

	if err := CopyDirectory(tempDir); err != nil {
		return fmt.Errorf("Couldn't copy template: %w", err)
	}

	if err := os.Mkdir(filepath.Join(tempDir, "ppt/media"), 0755); err != nil {
		return err
	}

	nslide, err := extractImages(src, filepath.Join(tempDir, "ppt/media"))
	if err != nil {
		return fmt.Errorf("Couldn't extract images: %w", err)
	}

	if err := cpSlides(tempDir, nslide); err != nil {
		return fmt.Errorf("Nao pode exec cpSlides:%w", err)
	}
	if err := genPresentationXML(tempDir, nslide); err != nil {
		return fmt.Errorf("Nao pode exec presentation.xml:%w", err)
	}
	if err := genContentTypesXML(tempDir, nslide); err != nil {
		return fmt.Errorf("Nao pode exec contentTypes:%w", err)
	}
	if err := genPresentationXMLRels(tempDir, nslide); err != nil {
		return fmt.Errorf("Nao pode exec xml rels:%w", err)
	}
	if err := zipit(tempDir, dest); err != nil {
		return fmt.Errorf("Nao pode zipar: %w", err)
	}

	return nil
}

func genPresentationXMLRels(pname string, num int) error {
	var s string
	s += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	s += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
	s += "<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles\" Target=\"tableStyles.xml\"/>"
	s += "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>"
	for i := num - 1; i >= 0; i-- {
		id := i + 8
		s += fmt.Sprintf("<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide-%d.xml\"/>", id, i)
	}
	s += "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>"
	s += "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
	s += "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps\" Target=\"viewProps.xml\"/>"
	s += "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps\" Target=\"presProps.xml\"/>"
	s += "</Relationships>"

	f, err := os.Create(pname + "/ppt/_rels/presentation.xml.rels")
	if err != nil {
		return err
	}
	f.WriteString(s)
	f.Sync()
	f.Close()
	return nil
}

func genContentTypesXML(pname string, num int) error {
	var s string
	s += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	s += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
	s += "<Default Extension=\"png\" ContentType=\"image/png\"/>"
	s += "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"
	s += "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
	s += "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
	s += "<Default Extension=\"JPG\" ContentType=\"image/jpeg\"/>"
	s += "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>"
	s += "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>"
	s += "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
	for i := num - 1; i >= 0; i-- {
		s += fmt.Sprintf("<Override PartName=\"/ppt/slides/slide-%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>", i)
	}
	s += "<Override PartName=\"/ppt/presProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presProps+xml\"/>"
	s += "<Override PartName=\"/ppt/viewProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml\"/>"
	s += "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
	s += "<Override PartName=\"/ppt/tableStyles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml\"/>"
	s += "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>"
	s += "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
	s += "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
	s += "</Types>"

	f, err := os.Create(pname + "/[Content_Types].xml")
	if err != nil {
		return err
	}
	f.WriteString(s)
	f.Sync()
	f.Close()
	return nil
}

func genPresentationXML(pname string, num int) error {
	read, err := ioutil.ReadFile(pname + "/ppt/presentation.xml")
	if err != nil {
		return err
	}
	newContents := string(read)

	for i := num - 1; i >= 0; i-- {
		sid := i + 256
		id := i + 8
		newContents = strings.Replace(newContents, "<p:sldIdLst>", fmt.Sprintf("<p:sldIdLst><p:sldId id=\"%d\" r:id=\"rId%d\"/>", sid, id), -1)
	}

	err = ioutil.WriteFile(pname+"/ppt/presentation.xml", []byte(newContents), 0)
	if err != nil {
		return err
	}
	return nil
}

func cpSlides(pname string, num int) error {
	for i := num - 1; i >= 0; i-- {
		Copy(pname+"/ppt/slides/slide1.xml", pname+fmt.Sprintf("/ppt/slides/slide-%d.xml", i))
		f, err := os.Create(pname + fmt.Sprintf("/ppt/slides/_rels/slide-%d.xml.rels", i))
		if err != nil {
			return err
		}
		f.WriteString(makeSlideXML(i))
		f.Sync()
		f.Close()
	}

	return nil
}

func makeSlideXML(num int) string {
	var s string
	s += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	s += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
	s += fmt.Sprintf("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/slide-%d.png\"/>", num)
	s += "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>"
	s += "</Relationships>"
	return s
}

func zipit(source, target string) error {
	zipfile, err := os.Create(target)
	if err != nil {
		return err
	}
	defer zipfile.Close()

	archive := zip.NewWriter(zipfile)
	defer archive.Close()

	base := filepath.Base(source)
	log.Printf("Base: %s", base)

	err = filepath.Walk(source, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			log.Printf("walking dir: %s", path)
			if source == path {
				return nil
			}
			path += "/"
		}

		header, err := zip.FileInfoHeader(info)
		if err != nil {
			return err
		}
		header.Name = path[len(source)+1:]
		header.Method = zip.Deflate

		writer, err := archive.CreateHeader(header)
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		file, err := os.Open(path)
		if err != nil {
			return err
		}
		defer file.Close()
		_, err = io.Copy(writer, file)
		return err
	})
	if err != nil {
		return err
	}
	if err = archive.Flush(); err != nil {
		return err
	}
	return nil
}

func extractImages(docu, path string) (int, error) {
	doc, err := fitz.New(docu)
	if err != nil {
		panic(err)
	}

	defer doc.Close()

	// Extract pages as images
	for n := 0; n < doc.NumPage(); n++ {
		img, err := doc.Image(n)
		if err != nil {
			return -1, fmt.Errorf("Nao pode ober imagem do pdf: %w", err)
		}

		f, err := os.Create(filepath.Join(path, fmt.Sprintf("slide-%d.png", n)))
		if err != nil {
			return -1, fmt.Errorf("Nao pode criar slide-%d.png: %w", n, err)
		}

		err = png.Encode(f, img)
		if err != nil {
			return -1, fmt.Errorf("Nao pode codificar png: %w", err)
		}

		f.Close()
	}
	return doc.NumPage(), nil
}
