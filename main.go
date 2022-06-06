package main

import (
	"fmt"
	"log"
	"reflect"
	"strconv"
	"time"

	"github.com/araddon/dateparse"
	"github.com/xuri/excelize/v2"
)

type Person struct {
	Name      string    `column:"A"`
	Age       int64     `column:"B"`
	Weight    float64   `column:"C"`
	BirthDate time.Time `column:"D"`
}

type Teacher struct {
	Person

	Class   int64  `column:"E"`
	Subject string `column:"F"`
}

func main() {
	f, err := excelize.OpenFile("./sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// 2行目以降に対して
	for _, row := range rows[1:] {
		p := Person{}
		if err := unmarshalRows(row, &p); err != nil {
			log.Fatal(err)
		}

		fmt.Printf("%+v\n", p)
	}
}

// 構造体のタグを元にrowの値を割り当てる
func unmarshalRows(row []string, v any) error {
	var f func(reflect.Value) error
	f = func(v reflect.Value) error {
		for i := 0; i < v.NumField(); i++ {
			structField := v.Type().Field(i)
			field := v.Field(i)

			// 構造体が埋め込まれていた場合、再帰的に探索
			if structField.Anonymous {
				f(field)
				continue
			}

			// columnタグの値を取得
			col := structField.Tag.Get("column")
			if col == "" {
				continue
			}

			str, err := getCol(row, col)
			if err != nil {
				return err
			}

			var val any

			// stringをフィールドの型に合うように変換
			switch field.Interface().(type) {
			case string:
				val = str
			case int64:
				val, err = strconv.ParseInt(str, 10, 64)
			case float64:
				val, err = strconv.ParseFloat(str, 64)
			case time.Time:
				val, err = dateparse.ParseAny(str)
			default:
				return fmt.Errorf("unexpected type: %s", field.Type())
			}
			if err != nil {
				log.Println(err)
				continue
			}

			// 構造体のフィールドに値をセット
			field.Set(reflect.ValueOf(val))
		}

		return nil
	}

	if err := f(reflect.ValueOf(v).Elem()); err != nil {
		return err
	}

	return nil
}

// 行から特定の列の値を取得
func getCol(row []string, col string) (string, error) {
	c, err := excelize.ColumnNameToNumber(col)
	if err != nil {
		return "", err
	}
	// 行の最後の空欄はGetRowsで読み込まれないため、そこを参照している場合は空文字を返す
	if len(row) < c {
		return "", nil
	}
	return row[c-1], nil
}
