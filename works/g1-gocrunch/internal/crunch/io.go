package crunch

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// ReadTable は拡張子(csv/tsv/json)から形式を判定して読み込む。
// JSONは「オブジェクトの配列」を想定し、キーの和集合を列にする。
func ReadTable(path string) (*Table, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".csv":
		return readDelimited(path, ',')
	case ".tsv":
		return readDelimited(path, '\t')
	case ".json":
		return readJSON(path)
	default:
		return nil, fmt.Errorf("%s: 対応していない形式です(csv/tsv/json)", path)
	}
}

func readDelimited(path string, sep rune) (*Table, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("開けません: %w", err)
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.Comma = sep
	r.FieldsPerRecord = -1 // 行ごとの列数ゆらぎは読み込み後に正規化する
	records, err := r.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("%s: 解析エラー: %w", path, err)
	}
	if len(records) == 0 {
		return &Table{}, nil
	}

	header := records[0]
	rows := make([][]string, 0, len(records)-1)
	for _, rec := range records[1:] {
		// 列数を見出しに合わせる(不足は空文字・超過は捨てず連結)
		row := make([]string, len(header))
		for i := range header {
			if i < len(rec) {
				row[i] = rec[i]
			}
		}
		if len(rec) > len(header) && len(header) > 0 {
			row[len(header)-1] += strings.Join(rec[len(header):], " ")
		}
		rows = append(rows, row)
	}
	return &Table{Header: header, Rows: rows}, nil
}

func readJSON(path string) (*Table, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("開けません: %w", err)
	}
	var objs []map[string]any
	if err := json.Unmarshal(data, &objs); err != nil {
		return nil, fmt.Errorf("%s: JSONは「オブジェクトの配列」形式にしてください: %w", path, err)
	}

	keySet := map[string]bool{}
	for _, o := range objs {
		for k := range o {
			keySet[k] = true
		}
	}
	header := make([]string, 0, len(keySet))
	for k := range keySet {
		header = append(header, k)
	}
	sort.Strings(header)

	rows := make([][]string, 0, len(objs))
	for _, o := range objs {
		row := make([]string, len(header))
		for i, k := range header {
			if v, ok := o[k]; ok {
				row[i] = fmt.Sprintf("%v", v)
			}
		}
		rows = append(rows, row)
	}
	return &Table{Header: header, Rows: rows}, nil
}

// WriteTable は拡張子から形式を判定して書き出す。
func WriteTable(t *Table, path string) error {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".csv":
		return writeDelimited(t, path, ',')
	case ".tsv":
		return writeDelimited(t, path, '\t')
	case ".json":
		return writeJSON(t, path)
	default:
		return fmt.Errorf("%s: 対応していない出力形式です(csv/tsv/json)", path)
	}
}

func writeDelimited(t *Table, path string, sep rune) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	w.Comma = sep
	if err := w.Write(t.Header); err != nil {
		return err
	}
	if err := w.WriteAll(t.Rows); err != nil {
		return err
	}
	w.Flush()
	return w.Error()
}

func writeJSON(t *Table, path string) error {
	objs := make([]map[string]string, 0, len(t.Rows))
	for _, row := range t.Rows {
		o := make(map[string]string, len(t.Header))
		for i, h := range t.Header {
			if i < len(row) {
				o[h] = row[i]
			}
		}
		objs = append(objs, o)
	}
	data, err := json.MarshalIndent(objs, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}
