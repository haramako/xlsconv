package main

import (
	"crypto/md5"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx"
)

var optOutput = flag.String("o", "", "output file")

func main() {
	flag.Parse()
	if flag.NArg() < 2 {
		fmt.Println("not input file")
		os.Exit(1)
	}

	args := flag.Args()
	file := args[0]
	sheet := args[1]

	defer func() {
		err := recover()
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
	}()

	var out io.Writer
	if *optOutput != "" {
		var err error
		out, err = os.OpenFile(*optOutput, os.O_CREATE|os.O_WRONLY, 0666)
		if err != nil {
			panic("cannot open file " + *optOutput)
		}
	} else {
		out = os.Stdout
	}

	hash, _, err := convert(file)
	if err != nil {
		panic(err)
	}

	data, err := ioutil.ReadFile(cachedSheetPath(hash, sheet))
	if err != nil {
		panic(err)
	}
	_, err = out.Write(data)
	if err != nil {
		panic(err)
	}
}

var tempDir string

func tempPath(name string) string {
	if tempDir == "" {
		tempDir = filepath.Join(os.TempDir(), "xls2tsv-go")
		os.MkdirAll(tempDir, 0777)
	}
	return filepath.Join(tempDir, name)
}

func cachedIndexPath(hash string) string {
	return tempPath(hash + ".index")
}

func cachedSheetPath(hash, sheet string) string {
	return tempPath(hash + "_" + sheet + ".tsv")
}

func fileHash(file string) (string, error) {
	data, err := ioutil.ReadFile(file)
	if err != nil {
		return "", err
	}

	sum := md5.Sum(data)
	return fmt.Sprintf("%x", sum), nil
}

func convert(file string) (string, []string, error) {
	hash, err := fileHash(file)
	if err != nil {
		return "", nil, err
	}

	_, err = os.Stat(cachedIndexPath(hash))
	if !os.IsNotExist(err) {
		// キャッシュがあった
		sheetsStr, err := ioutil.ReadFile(cachedIndexPath(hash))
		if err != nil {
			return "", nil, err
		}
		sheets := strings.Split(string(sheetsStr), "\n")
		return hash, sheets, nil
	} else {
		// キャッシュがなかった
		sheets, err := convertCache(file, hash)
		if err != nil {
			return "", nil, err
		}
		err = ioutil.WriteFile(cachedIndexPath(hash), []byte(strings.Join(sheets, "\n")), 0666)
		if err != nil {
			return "", nil, err
		}
		return hash, sheets, nil
	}
}

func convertCache(file string, hash string) ([]string, error) {
	fmt.Fprintf(os.Stderr, "converting file %v\n", file)

	sheets := []string{}

	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return nil, fmt.Errorf("can't open file %s", file)
	}

	for _, sheet := range xls.Sheets {
		fmt.Fprintf(os.Stderr, "converting sheet %v\n", sheet.Name)

		out, err := os.OpenFile(cachedSheetPath(hash, sheet.Name), os.O_CREATE|os.O_WRONLY, 0666)
		if err != nil {
			return nil, err
		}
		defer func() {
			out.Close()
		}()

		sheets = append(sheets, sheet.Name)

		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				str := cell.String()
				str = strings.Replace(str, "\\", "\\\\", -1)
				str = strings.Replace(str, "\n", "\\n", -1)
				str = strings.Replace(str, "\t", "\\t", -1)
				fmt.Fprintf(out, "%s\t", str)
			}
			fmt.Fprintf(out, "\n")
		}
	}

	return sheets, nil
}
