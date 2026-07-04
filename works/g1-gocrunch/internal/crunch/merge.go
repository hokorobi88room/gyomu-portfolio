package crunch

import (
	"fmt"
	"runtime"
	"sort"
	"sync"
)

// MergeResult は結合結果と、失敗ファイルの一覧(黙って捨てない)。
type MergeResult struct {
	Table    *Table
	Failed   map[string]error // パス → エラー
	NumFiles int
}

// MergeFiles は複数ファイルを並行読み込みして1つの表に結合する。
// 列名の和集合を取り、ファイルごとの列順の違いは列名で揃える。
// withSource が真なら「元ファイル」列を先頭に追加する。
func MergeFiles(paths []string, workers int, withSource bool) (*MergeResult, error) {
	if len(paths) == 0 {
		return nil, fmt.Errorf("対象ファイルがありません")
	}
	if workers <= 0 {
		workers = runtime.NumCPU()
	}

	type item struct {
		path  string
		table *Table
		err   error
	}

	// 並行読み込み(ワーカープール)。結果の順序はパス順に整列し直す。
	jobs := make(chan string)
	results := make(chan item)
	var wg sync.WaitGroup
	for w := 0; w < workers; w++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for p := range jobs {
				t, err := ReadTable(p)
				results <- item{path: p, table: t, err: err}
			}
		}()
	}
	go func() {
		for _, p := range paths {
			jobs <- p
		}
		close(jobs)
		wg.Wait()
		close(results)
	}()

	tables := map[string]*Table{}
	failed := map[string]error{}
	for it := range results {
		if it.err != nil {
			failed[it.path] = it.err
		} else {
			tables[it.path] = it.table
		}
	}
	if len(tables) == 0 {
		return nil, fmt.Errorf("すべてのファイルの読み込みに失敗しました(%d件)", len(failed))
	}

	// 列の和集合(初出順を保つため、パス順に走査)
	sorted := make([]string, 0, len(tables))
	for _, p := range paths {
		if _, ok := tables[p]; ok {
			sorted = append(sorted, p)
		}
	}
	var header []string
	seen := map[string]bool{}
	for _, p := range sorted {
		for _, h := range tables[p].Header {
			if !seen[h] {
				seen[h] = true
				header = append(header, h)
			}
		}
	}

	outHeader := header
	if withSource {
		outHeader = append([]string{"元ファイル"}, header...)
	}

	total := 0
	for _, t := range tables {
		total += len(t.Rows)
	}
	rows := make([][]string, 0, total)
	for _, p := range sorted {
		t := tables[p]
		// 列名 → この表での添字
		idx := make(map[string]int, len(t.Header))
		for i, h := range t.Header {
			idx[h] = i
		}
		for _, src := range t.Rows {
			row := make([]string, len(outHeader))
			off := 0
			if withSource {
				row[0] = p
				off = 1
			}
			for i, h := range header {
				if j, ok := idx[h]; ok && j < len(src) {
					row[i+off] = src[j]
				}
			}
			rows = append(rows, row)
		}
	}

	return &MergeResult{
		Table:    &Table{Header: outHeader, Rows: rows},
		Failed:   failed,
		NumFiles: len(tables),
	}, nil
}

// GroupSum は groupCol ごとに sumCol を合計する(降順)。
// 数値にできないセルは0ではなくスキップし、件数を返す(黙って0にしない)。
func GroupSum(t *Table, groupCol, sumCol string) (*Table, int, error) {
	gi, err := t.ColumnIndex(groupCol)
	if err != nil {
		return nil, 0, err
	}
	si, err := t.ColumnIndex(sumCol)
	if err != nil {
		return nil, 0, err
	}

	sums := map[string]float64{}
	counts := map[string]int{}
	skipped := 0
	for _, row := range t.Rows {
		if gi >= len(row) || si >= len(row) {
			skipped++
			continue
		}
		v, err := parseNumber(row[si])
		if err != nil {
			skipped++
			continue
		}
		sums[row[gi]] += v
		counts[row[gi]]++
	}

	keys := make([]string, 0, len(sums))
	for k := range sums {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(a, b int) bool { return sums[keys[a]] > sums[keys[b]] })

	rows := make([][]string, 0, len(keys))
	for _, k := range keys {
		rows = append(rows, []string{k, formatNumber(sums[k]), fmt.Sprintf("%d", counts[k])})
	}
	return &Table{
		Header: []string{groupCol, sumCol + "(合計)", "件数"},
		Rows:   rows,
	}, skipped, nil
}
