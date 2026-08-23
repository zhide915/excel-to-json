package main

import (
	"os"

	"github.com/zhide915/excel-to-json/internal/cli"
)

func main() {
	os.Exit(cli.Main(os.Args[1:]))
}
