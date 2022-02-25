package main

import (
	"fmt"
	"reflect"
	"strconv"
	"time"

	"github.com/araddon/dateparse"
	"github.com/xuri/excelize/v2"
)

type Person struct {
	Name      string    `json:"name" column:"A"`
	Age       int64     `json:"age" column:"B"`
	Weight    float64   `json:"weight" column:"C"`
	Birthdate time.Time `json:"birth_date" column:"D"`
}

func main() {
	f, err := excelize.OpenFile("./sample.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 2行目以降に対して
	for _, row := range rows[1:] {
		p := Person{}
		if err := setValue(row, &p); err != nil {
			fmt.Println(err)
			return
		}

		fmt.Printf("%+v\n", p)
	}
}

// 構造体のタグを元にrowの値を割り当てる
func setValue(row []string, in interface{}) error {
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

			var val interface{}

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
				return fmt.Errorf("unknown type: %s", field.Type().Name())
			}
			if err != nil {
				return err
			}

			// 構造体のフィールドに値をセット
			field.Set(reflect.ValueOf(val))
		}

		return nil
	}

	if err := f(reflect.ValueOf(in).Elem()); err != nil {
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
