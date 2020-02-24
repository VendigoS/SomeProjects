package main

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"os"
)

func dirTree(out *bytes.Buffer, currentPath string, full bool) error{
	var fileSize string //for full output. Show file.Size()
	var prefix string
	
	files, err := ioutil.ReadDir(currentPath)
	if err != nil {
		return err
	}
	
	onlyDirs := 0 //these two vars for counting directories among all files for normal working with tabs and lines
	indexDir := 0
	
	//sum amount of dirs
	if !full {
		for i := 0; i < len(files); i++ {
			if files[i].IsDir() {
				onlyDirs++
			}
		}
	}
	
	for index, file := range files {
		//display file's size
		fileSize = ""
		if file.Mode().IsRegular() && full {
			if file.Size() > 0 {
				fileSize = " ("+ fmt.Sprint(file.Size()) + "b)"
			} else {
				fileSize = " (empty)"	
			}
		}
		if file.IsDir() {
			indexDir++
		}
		//check if we should display file info
		if file.IsDir() || file.Mode().IsRegular() && full {
			os.Stdout.Write(out.Bytes())
			//if it is the last dir/file in list then display truncated line
			if index == len(files) - 1 && full || indexDir == onlyDirs  && !full {
				os.Stdout.Write([]byte(`└───`))
				os.Stdout.Write([]byte("" + file.Name() + fileSize + "\n"))
				prefix = "\t"
			} else if index < len(files) - 1 && full || indexDir < onlyDirs && !full {
				os.Stdout.Write([]byte(`├───`))
				os.Stdout.Write([]byte("" + file.Name() + fileSize + "\n"))
				prefix = "│\t"
			}
			
		}
		//go recursive into another dir
		if file.IsDir() {
			bufStr := out.String()		
			out.Write([]byte(prefix))
			dirTree(out, currentPath + string(os.PathSeparator) + file.Name(), full)
			out.Reset() //clear buffer and full with old value of prefix
			out.Write([]byte(bufStr))
		}
	
	}

	
	/*
	//Trying to use way without recurse, but filepath.Walk doesn't provide info about full path :(
	err := filepath.Walk(path, func(inPath string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		
		if info.IsDir() || info.IsDir() != true && full {
			if full {
				fileSize = " ("+ fmt.Sprint(info.Size()) + "b)"
			}
			fmt.Println(info.Name() + fileSize)
		}
		return nil
	})
	*/
	return err
}

const usageText = `Usage of program:
go run main.go . [-f]`

func main(){
	var currentPath string
	var full bool
	var err error

	/*
	//There was trying to use flag module, but it doesn't support non-positional arguments before positional
	flag.BoolVar(&full, "f", false, "Complete pass with files or only dirs")
	flag.Parse()
	
	if len(flag.Args()) == 0 {
		fmt.Println(fmt.Errorf("Missing non-positional arguments. %s", usageText))
		os.Exit(1)
	}
	*/

	switch(len(os.Args)) {
		case 3:
			if os.Args[2] == `-f` {
				full = true
			} else {
				fmt.Println(fmt.Errorf("Unknown argument %s. %s",os.Args[2], usageText))
	                        os.Exit(1)
			}
			currentPath = os.Args[1]
		case 2:
			full = false
			currentPath = os.Args[1]
		case 1:	
			fmt.Println(fmt.Errorf("Missing arguments. %s", usageText))
			os.Exit(1)
		default:
			fmt.Println(fmt.Errorf("Too much arguments. %s", usageText))
			os.Exit(1)
	}
	
	out := new(bytes.Buffer)
	err = dirTree(out, currentPath, full)
	if err != nil {
		os.Exit(1)
	} else {
		//os.Stdout.Write(out.Bytes())
	}	
}
