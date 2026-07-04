package crunch

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

func writeFile(t testing.TB, dir, name, content string) string {
	t.Helper()
	p := filepath.Join(dir, name)
	if err := os.WriteFile(p, []byte(content), 0o644); err != nil {
		t.Fatal(err)
	}
	return p
}

// ---------- 読み書き ----------

func TestReadTableFormats(t *testing.T) {
	dir := t.TempDir()
	tests := []struct {
		name    string
		file    string
		content string
		wantHdr []string
		wantN   int
	}{
		{"csv", "a.csv", "支店,売上\n東京,100\n大阪,200\n", []string{"支店", "売上"}, 2},
		{"tsv", "a.tsv", "支店\t売上\n東京\t100\n", []string{"支店", "売上"}, 1},
		{"json", "a.json", `[{"支店":"東京","売上":100}]`, []string{"売上", "支店"}, 1},
		{"列数ゆらぎ", "b.csv", "a,b\n1\n1,2,3\n", []string{"a", "b"}, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tbl, err := ReadTable(writeFile(t, dir, tt.file, tt.content))
			if err != nil {
				t.Fatal(err)
			}
			if len(tbl.Rows) != tt.wantN {
				t.Errorf("行数 = %d, want %d", len(tbl.Rows), tt.wantN)
			}
			if fmt.Sprint(tbl.Header) != fmt.Sprint(tt.wantHdr) {
				t.Errorf("header = %v, want %v", tbl.Header, tt.wantHdr)
			}
		})
	}
}

func TestReadTableErrors(t *testing.T) {
	dir := t.TempDir()
	tests := []struct {
		name string
		path string
	}{
		{"存在しない", filepath.Join(dir, "nai.csv")},
		{"未対応拡張子", writeFile(t, dir, "a.xlsx", "x")},
		{"壊れたJSON", writeFile(t, dir, "bad.json", `{"not":"array"}`)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if _, err := ReadTable(tt.path); err == nil {
				t.Error("エラーになるべき入力でnilが返った")
			}
		})
	}
}

func TestConvertRoundtrip(t *testing.T) {
	dir := t.TempDir()
	src := writeFile(t, dir, "s.csv", "名前,数\nあ,1\nい,2\n")
	tbl, err := ReadTable(src)
	if err != nil {
		t.Fatal(err)
	}
	jsonPath := filepath.Join(dir, "out.json")
	if err := WriteTable(tbl, jsonPath); err != nil {
		t.Fatal(err)
	}
	back, err := ReadTable(jsonPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(back.Rows) != 2 {
		t.Errorf("JSON往復で行数が変わった: %d", len(back.Rows))
	}
}

// ---------- 結合 ----------

func TestMergeAlignsColumnsByName(t *testing.T) {
	dir := t.TempDir()
	// 列順が違う2ファイル+片方にしかない列
	p1 := writeFile(t, dir, "1.csv", "支店,売上\n東京,100\n")
	p2 := writeFile(t, dir, "2.csv", "売上,支店,担当\n200,大阪,佐藤\n")

	res, err := MergeFiles([]string{p1, p2}, 2, true)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"元ファイル", "支店", "売上", "担当"}
	if fmt.Sprint(res.Table.Header) != fmt.Sprint(want) {
		t.Fatalf("header = %v, want %v", res.Table.Header, want)
	}
	if len(res.Table.Rows) != 2 {
		t.Fatalf("行数 = %d, want 2", len(res.Table.Rows))
	}
	// 大阪行: 列名で正しく整列していること
	osaka := res.Table.Rows[1]
	if osaka[1] != "大阪" || osaka[2] != "200" || osaka[3] != "佐藤" {
		t.Errorf("列整列が不正: %v", osaka)
	}
}

func TestMergeContinuesOnBrokenFile(t *testing.T) {
	dir := t.TempDir()
	ok := writeFile(t, dir, "ok.csv", "a\n1\n")
	bad := writeFile(t, dir, "bad.json", "{broken")

	res, err := MergeFiles([]string{ok, bad}, 2, false)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Failed) != 1 {
		t.Errorf("失敗ファイルが記録されていない: %v", res.Failed)
	}
	if res.NumFiles != 1 || len(res.Table.Rows) != 1 {
		t.Errorf("正常ファイルの処理が継続していない")
	}
}

func TestMergeAllBrokenReturnsError(t *testing.T) {
	dir := t.TempDir()
	bad := writeFile(t, dir, "bad.json", "{broken")
	if _, err := MergeFiles([]string{bad}, 1, false); err == nil {
		t.Error("全滅時はエラーを返すべき")
	}
}

// ---------- 集計 ----------

func TestGroupSum(t *testing.T) {
	tbl := &Table{
		Header: []string{"支店", "売上"},
		Rows: [][]string{
			{"東京", "1,000"},   // カンマ入り
			{"東京", "¥500"},    // 通貨記号入り
			{"大阪", "2000"},
			{"大阪", "数値じゃない"}, // スキップされるべき
		},
	}
	st, skipped, err := GroupSum(tbl, "支店", "売上")
	if err != nil {
		t.Fatal(err)
	}
	if skipped != 1 {
		t.Errorf("skipped = %d, want 1", skipped)
	}
	// 降順: 大阪2000 → 東京1500
	if st.Rows[0][0] != "大阪" || st.Rows[0][1] != "2000" {
		t.Errorf("1位が不正: %v", st.Rows[0])
	}
	if st.Rows[1][0] != "東京" || st.Rows[1][1] != "1500" {
		t.Errorf("2位が不正: %v", st.Rows[1])
	}
}

func TestGroupSumMissingColumn(t *testing.T) {
	tbl := &Table{Header: []string{"a"}, Rows: [][]string{{"1"}}}
	if _, _, err := GroupSum(tbl, "支店", "売上"); err == nil {
		t.Error("存在しない列はエラーになるべき")
	}
}

// ---------- ベンチマーク(README掲載の実測値の根拠) ----------

func makeBenchFiles(b *testing.B, n, rowsPer int) []string {
	b.Helper()
	dir := b.TempDir()
	paths := make([]string, n)
	for i := 0; i < n; i++ {
		content := "支店,商品,売上\n"
		for r := 0; r < rowsPer; r++ {
			content += fmt.Sprintf("支店%02d,商品%d,%d\n", i%20, r%10, 1000+r)
		}
		paths[i] = writeFile(b, dir, fmt.Sprintf("f%04d.csv", i), content)
	}
	return paths
}

func BenchmarkMerge2000Files(b *testing.B) {
	paths := makeBenchFiles(b, 2000, 50)
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		res, err := MergeFiles(paths, 0, false)
		if err != nil {
			b.Fatal(err)
		}
		if len(res.Table.Rows) != 2000*50 {
			b.Fatalf("行数不正: %d", len(res.Table.Rows))
		}
	}
}
