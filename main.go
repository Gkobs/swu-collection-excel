package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
	"sort"
	"strconv"
	//"os"
	//"github.com/xuri/excelize/v2"
)

type Data struct {
	TotalCards int    `json:"total_cards"`
	Data       []Card `json:"data"`
}

type Card struct {
	Set          string   `json:"Set"`
	Number       string   `json:"Number"`
	Name         string   `json:"Name"`
	Type         string   `json:"Type"`
	Aspects      []string `json:"Aspects"`
	Traits       []string `json:"Traits"`
	Arenas       []string `json:"Arenas"`
	Cost         string   `json:"Cost"`
	Power        string   `json:"Power"`
	HP           string   `json:"HP"`
	FrontText    string   `json:"FrontText"`
	DoubleSided  bool     `json:"DoubleSided"`
	Rarity       string   `json:"Rarity"`
	Unique       bool     `json:"Unique"`
	Keywords     []string `json:"Keywords"`
	Artist       string   `json:"Artist"`
	VariantType  string   `json:"VariantType"`
	MarketPrice  string   `json:"MarketPrice"`
	FoilPrice    string   `json:"FoilPrice"`
	FrontArt     string   `json:"FrontArt"`
	LowFoilPrice string   `json:"LowFoilPrice"`
	LowPrice     string   `json:"LowPrice"`
}

func main() {

	set := "lof"
	d, err := http.Get(fmt.Sprintf("https://api.swu-db.com/cards/%v", set))
	if err != nil {
		log.Fatal(err)
	}

	defer d.Body.Close()
	bytes, err := io.ReadAll(d.Body)
	if err != nil {
		log.Fatal(err)
	}

	var data Data
	err = json.Unmarshal(bytes, &data)
	if err != nil {
		log.Fatal(err)
	}
	cards := data.Data

	sort.Slice(cards, func(i, j int) bool {
		num1, err := strconv.Atoi(cards[i].Number)
		num2, err := strconv.Atoi(cards[j].Number)
		if err != nil {
			log.Fatal(err)
		}

		return num1 < num2
	})
	f := excelize.NewFile()
	header := [6]string{"Number", "Name", "Rarity", "Type", "Card Count", "Completion"}
	for i, h := range header {
		err = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+i)), 1), h)

		if err != nil {
			log.Fatal(err)
		}
	}

	indexLimit := 263
	rowOffset := 2
	for i, card := range cards {
		if i > indexLimit {
			break
		}

		row := i + rowOffset
		err = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65)), row), card.Number)
		err = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+1)), row), card.Name)
		err = f.SetCellHyperLink("Sheet1", fmt.Sprintf("%s%d", string(rune(65+1)), row), card.FrontArt, "External", excelize.HyperlinkOpts{Display: &card.Name, Tooltip: &card.FrontArt})
		err = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+2)), row), card.Rarity)
		err = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+3)), row), card.Type)
		err := f.SetCellFormula("Sheet1", fmt.Sprintf("%s%d", string(rune(65+5)), row), fmt.Sprintf("IF(OR(%s = \"Leader\", %s = \"Base\"), IF(%s >= 1, true, false), IF(%s >= 3, true, false))", fmt.Sprintf("%s%d", string(rune(65+3)), row), fmt.Sprintf("%s%d", string(rune(65+3)), row), fmt.Sprintf("%s%d", string(rune(65+4)), row), fmt.Sprintf("%s%d", string(rune(65+4)), row)))
		if err != nil {
			log.Fatal(err)
		}
	}

	//dir, err := os.Getwd()
	//if err != nil {
	//	log.Fatal(err)
	//}

	err = f.SaveAs(fmt.Sprintf("./files/%s.xlsx", set))
	if err != nil {
		log.Fatal(err)
	}
}
