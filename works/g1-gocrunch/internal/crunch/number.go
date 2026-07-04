package crunch

import (
	"fmt"
	"strconv"
	"strings"
)

// parseNumber は "1,234" "¥5,000" " 12.5 " のような現場のセル値を数値にする。
func parseNumber(s string) (float64, error) {
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, ",", "")
	s = strings.TrimPrefix(s, "¥")
	s = strings.TrimPrefix(s, "￥")
	s = strings.TrimSuffix(s, "円")
	if s == "" {
		return 0, fmt.Errorf("空文字")
	}
	return strconv.ParseFloat(s, 64)
}

// formatNumber は合計値を表示用に整形する(整数なら小数点を出さない)。
func formatNumber(v float64) string {
	if v == float64(int64(v)) {
		return strconv.FormatInt(int64(v), 10)
	}
	return strconv.FormatFloat(v, 'f', 2, 64)
}
