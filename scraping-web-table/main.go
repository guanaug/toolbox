package main

import (
	"fmt"
	"github.com/PuerkitoBio/goquery"
	"github.com/gocolly/colly/v2"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
)

var url string
var output = excelize.NewFile()
var sheetNum = 0

func usage() {
	fmt.Printf(`
Usage:
	%s $URL
Example:
	%s www.baidu.com
`, os.Args[0], os.Args[0])
}

func init() {
	if len(os.Args) < 2 {
		usage()
	}

	url = os.Args[1]
}

func main() {
	c := colly.NewCollector()

	// Find and visit all links
	c.OnHTML("table", func(e *colly.HTMLElement) {
		processTable(e)
	})

	c.OnRequest(func(r *colly.Request) {
		fmt.Println("正在爬取网页...", url)
	})

	c.OnResponse(func(r *colly.Response) {
		fmt.Println("爬取网页成功，开始解析数据...")
	})

	if err := c.Visit(url); err != nil {
		fmt.Println("爬取网页失败", err)
	}

	//output.NewSheet("Sheet1")
	//rows := []string{"a", "b", "c"}
	//if err := output.SetSheetRow("Sheet1", fmt.Sprintf("A%d", 1), &rows); err != nil {
	//	fmt.Println(err)
	//	return
	//}

	fmt.Println("解析数据完成 ...")
	if err := output.SaveAs("table.xlsx"); err != nil {
		fmt.Println("保存文件失败，请重试")
	}
}

func processTable(e *colly.HTMLElement) {
	sheetNum++
	sheetName := "Sheet" + strconv.Itoa(sheetNum)
	output.NewSheet(sheetName)

	idx := 1
	idx = parseHead(sheetName, idx, e)
	idx = parseBody(sheetName, idx, e)
	parseFoot(sheetName, idx, e)
}

func parseHead(sheetName string, idx int, e *colly.HTMLElement) int {
	tbody := e.DOM.Find("thead")
	tr := tbody.Find("tr")
	tr.Each(func(i int, selection *goquery.Selection) {
		td := selection.Find("th")
		cols := make([]string, 0, td.Size())

		td.Each(func(i int, selection *goquery.Selection) {
			cols = append(cols, selection.Text())
		})

		if err := output.SetSheetRow(sheetName, fmt.Sprintf("A%d", idx), &cols); err != nil {
			fmt.Printf("解析 %s 第 %d 行失败: %+v\n", sheetName, i, err)
			return
		}
		idx++
	})

	return idx
}

func parseBody(sheetName string, idx int, e *colly.HTMLElement) int {
	tbody := e.DOM.Find("tbody")
	tr := tbody.Find("tr")
	tr.Each(func(i int, selection *goquery.Selection) {
		td := selection.Find("td")
		cols := make([]string, 0, td.Size())

		td.Each(func(i int, selection *goquery.Selection) {
			cols = append(cols, selection.Text())
		})

		if err := output.SetSheetRow(sheetName, fmt.Sprintf("A%d", idx), &cols); err != nil {
			fmt.Printf("解析 %s 第 %d 行失败: %+v\n", sheetName, i, err)
			return
		}
		idx++
	})

	return idx
}

func parseFoot(sheetName string, idx int, e *colly.HTMLElement) int {
	tbody := e.DOM.Find("tfoot")
	tr := tbody.Find("tr")
	tr.Each(func(i int, selection *goquery.Selection) {
		td := selection.Find("td")
		cols := make([]string, 0, td.Size())

		td.Each(func(i int, selection *goquery.Selection) {
			cols = append(cols, selection.Text())
		})

		if err := output.SetSheetRow(sheetName, fmt.Sprintf("A%d", idx), &cols); err != nil {
			fmt.Printf("解析 %s 第 %d 行失败: %+v\n", sheetName, i, err)
			return
		}
		idx++
	})

	return idx
}
