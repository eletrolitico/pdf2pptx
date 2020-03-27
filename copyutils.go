package pptx

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/markbates/pkger"
)

// CopyDirectory copia dir inteiro
func CopyDirectory(dest string) error {
	return pkger.Walk("/pptxTemplate", func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		pp := strings.Split(path, ":")[1]
		if pp == "/pptxTemplate" {
			return nil
		}
		pp = strings.Replace(pp, "/pptxTemplate", "/", 1)
		if info.IsDir() {
			err := os.Mkdir(filepath.Join(dest, pp), info.Mode().Perm())
			if err != nil {
				return err
			}
		} else {
			f, err := pkger.Open(path)
			if err != nil {
				return fmt.Errorf("Couldn'd open %s: %w", path, err)
			}
			d, err := os.Create(filepath.Join(dest, pp))
			if err != nil {
				return err
			}
			if _, err := io.Copy(d, f); err != nil {
				return err
			}
			if err := f.Close(); err != nil {
				return err
			}
			if err := d.Sync(); err != nil {
				return err
			}
			if err := d.Close(); err != nil {
				return err
			}
		}
		return nil
	})
}

// Copy func de copia
func Copy(srcFile, dstFile string) error {
	out, err := os.Create(dstFile)
	if err != nil {
		return err
	}

	defer out.Close()

	in, err := os.Open(srcFile)
	defer in.Close()
	if err != nil {
		return err
	}

	_, err = io.Copy(out, in)
	if err != nil {
		return err
	}

	return nil
}
