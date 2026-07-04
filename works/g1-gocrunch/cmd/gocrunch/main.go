// gocrunch — 大量のCSV/TSV/JSONを高速に 結合(merge)・変換(convert)・集計(stats) する
// 依存ゼロの単一バイナリCLI。Excelでは開けない量のファイル処理を想定する。
//
// 使い方:
//	gocrunch merge  -in "data/*.csv" -out merged.csv [-source] [-workers 8] [-dry-run]
//	gocrunch convert -in file.csv -out file.json
//	gocrunch stats  -in "data/*.csv" -group 支店 -sum 売上 [-out stats.csv]
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/example/gocrunch/internal/crunch"
)

func main() {
	if len(os.Args) < 2 {
		usage()
		os.Exit(2)
	}
	var err error
	switch os.Args[1] {
	case "merge":
		err = cmdMerge(os.Args[2:])
	case "convert":
		err = cmdConvert(os.Args[2:])
	case "stats":
		err = cmdStats(os.Args[2:])
	case "-h", "--help", "help":
		usage()
		return
	default:
		fmt.Fprintf(os.Stderr, "不明なコマンド: %s\n\n", os.Args[1])
		usage()
		os.Exit(2)
	}
	if err != nil {
		fmt.Fprintln(os.Stderr, "エラー:", err)
		os.Exit(1)
	}
}

func usage() {
	fmt.Println(`gocrunch — 大量のCSV/TSV/JSONを高速処理する単一バイナリCLI

コマンド:
  merge    複数ファイルを1つに結合(列名で自動整列・並行読み込み)
           gocrunch merge -in "data/*.csv" -out merged.csv [-source] [-workers N] [-dry-run]
  convert  形式変換(csv/tsv/json は拡張子で自動判定)
           gocrunch convert -in file.csv -out file.json
  stats    グループ集計(合計・件数、降順)
           gocrunch stats -in "data/*.csv" -group 支店 -sum 売上 [-out stats.csv]`)
}

func cmdMerge(args []string) error {
	fs := flag.NewFlagSet("merge", flag.ExitOnError)
	in := fs.String("in", "", "入力(glob可: \"data/*.csv\")")
	out := fs.String("out", "merged.csv", "出力ファイル")
	source := fs.Bool("source", false, "「元ファイル」列を追加する")
	workers := fs.Int("workers", 0, "並行数(0=CPUコア数)")
	dryRun := fs.Bool("dry-run", false, "書き込まずに対象と件数だけ表示")
	_ = fs.Parse(args)

	paths, err := expand(*in)
	if err != nil {
		return err
	}
	fmt.Printf("対象: %d ファイル\n", len(paths))
	if *dryRun {
		for _, p := range paths {
			fmt.Println("  ", p)
		}
		fmt.Println("(dry-run: 書き込みなし)")
		return nil
	}

	t0 := time.Now()
	res, err := crunch.MergeFiles(paths, *workers, *source)
	if err != nil {
		return err
	}
	if err := crunch.WriteTable(res.Table, *out); err != nil {
		return err
	}

	fmt.Printf("完了: %d ファイル / %d 行 → %s(%.2f 秒)\n",
		res.NumFiles, len(res.Table.Rows), *out, time.Since(t0).Seconds())
	if len(res.Failed) > 0 {
		fmt.Printf("⚠ 読み込み失敗 %d 件(処理は継続しました):\n", len(res.Failed))
		for p, e := range res.Failed {
			fmt.Printf("   %s: %v\n", p, e)
		}
		return fmt.Errorf("%d 件のファイルが読めませんでした", len(res.Failed))
	}
	return nil
}

func cmdConvert(args []string) error {
	fs := flag.NewFlagSet("convert", flag.ExitOnError)
	in := fs.String("in", "", "入力ファイル")
	out := fs.String("out", "", "出力ファイル(拡張子で形式判定)")
	_ = fs.Parse(args)
	if *in == "" || *out == "" {
		return fmt.Errorf("-in と -out を指定してください")
	}

	t, err := crunch.ReadTable(*in)
	if err != nil {
		return err
	}
	if err := crunch.WriteTable(t, *out); err != nil {
		return err
	}
	fmt.Printf("完了: %s(%d 行)→ %s\n", *in, len(t.Rows), *out)
	return nil
}

func cmdStats(args []string) error {
	fs := flag.NewFlagSet("stats", flag.ExitOnError)
	in := fs.String("in", "", "入力(glob可)")
	group := fs.String("group", "", "グループ化する列名")
	sum := fs.String("sum", "", "合計する列名")
	out := fs.String("out", "", "出力ファイル(省略時は画面表示)")
	workers := fs.Int("workers", 0, "並行数(0=CPUコア数)")
	_ = fs.Parse(args)
	if *group == "" || *sum == "" {
		return fmt.Errorf("-group と -sum を指定してください")
	}

	paths, err := expand(*in)
	if err != nil {
		return err
	}
	res, err := crunch.MergeFiles(paths, *workers, false)
	if err != nil {
		return err
	}
	st, skipped, err := crunch.GroupSum(res.Table, *group, *sum)
	if err != nil {
		return err
	}

	if *out != "" {
		if err := crunch.WriteTable(st, *out); err != nil {
			return err
		}
		fmt.Printf("完了: %s へ出力(%d グループ)\n", *out, len(st.Rows))
	} else {
		fmt.Printf("%-24s %16s %8s\n", st.Header[0], st.Header[1], st.Header[2])
		for _, row := range st.Rows {
			fmt.Printf("%-24s %16s %8s\n", row[0], row[1], row[2])
		}
	}
	if skipped > 0 {
		fmt.Printf("⚠ 数値にできず集計から除外した行: %d(黙って0にはしません)\n", skipped)
	}
	return nil
}

func expand(pattern string) ([]string, error) {
	if pattern == "" {
		return nil, fmt.Errorf("-in を指定してください")
	}
	paths, err := filepath.Glob(pattern)
	if err != nil {
		return nil, fmt.Errorf("globが不正です: %w", err)
	}
	if len(paths) == 0 {
		return nil, fmt.Errorf("対象がありません: %s", pattern)
	}
	return paths, nil
}
