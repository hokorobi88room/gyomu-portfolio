// Package crunch は表形式データ(CSV/TSV/JSON)の読み書き・結合・集計を提供する。
// 「Excelでは開けない量のファイルを、依存ゼロの単一バイナリで処理する」ことが目的。
package crunch

import "fmt"

// Table は表形式データのメモリ表現。
type Table struct {
	Header []string
	Rows   [][]string
}

// ColumnIndex は列名から添字を返す。見つからなければエラー。
func (t *Table) ColumnIndex(name string) (int, error) {
	for i, h := range t.Header {
		if h == name {
			return i, nil
		}
	}
	return -1, fmt.Errorf("列 %q がありません(存在する列: %v)", name, t.Header)
}
