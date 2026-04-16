package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type result struct {
	Label      string  `json:"label"`
	Real       float64 `json:"real"`
	RealMin    float64 `json:"real_min"`
	RealStddev float64 `json:"real_stddev"`
	RSSDeltaKB uint64  `json:"rss_delta_kb"`
	Iterations int     `json:"iterations"`
	Size       int64   `json:"size,omitempty"`
}

var (
	rows       = intEnv("RBXL_BENCH_ROWS", 5000)
	cols       = intEnv("RBXL_BENCH_COLS", 10)
	warmup     = intEnv("RBXL_BENCH_WARMUP", 1)
	iterations = intEnv("RBXL_BENCH_ITERATIONS", 5)
	readPath   = os.Getenv("RBXL_BENCH_READ_PATH")
)

func main() {
	header, body := buildDataset(rows, cols)
	tmpDir, err := os.MkdirTemp("", "rbxl-go-bench-")
	if err != nil {
		fail(err)
	}
	defer os.RemoveAll(tmpDir)

	writePath := filepath.Join(tmpDir, "excelize.xlsx")
	results := make([]result, 0, 2)

	writeResult := measure("excelize write", func() error {
		return writeWithExcelize(writePath, header, body)
	})
	info, err := os.Stat(writePath)
	if err != nil {
		fail(err)
	}
	writeResult.Size = info.Size()
	results = append(results, writeResult)

	targetReadPath := writePath
	if readPath != "" {
		targetReadPath = readPath
	}
	results = append(results, measure("excelize read", func() error {
		_, err := readWithExcelize(targetReadPath)
		return err
	}))

	if err := json.NewEncoder(os.Stdout).Encode(results); err != nil {
		fail(err)
	}
}

func intEnv(name string, fallback int) int {
	if raw := os.Getenv(name); raw != "" {
		if parsed, err := strconv.Atoi(raw); err == nil {
			return parsed
		}
	}
	return fallback
}

func buildDataset(rowCount, colCount int) ([]any, [][]any) {
	header := make([]any, colCount)
	for i := 0; i < colCount; i++ {
		header[i] = fmt.Sprintf("col_%d", i+1)
	}

	body := make([][]any, rowCount)
	for row := 0; row < rowCount; row++ {
		values := make([]any, colCount)
		for col := 0; col < colCount; col++ {
			switch col % 4 {
			case 0:
				values[col] = row
			case 1:
				values[col] = fmt.Sprintf("row-%d-col-%d", row, col)
			case 2:
				values[col] = (row+col)%2 == 1
			default:
				values[col] = float64((row*100)+col) / 10.0
			}
		}
		body[row] = values
	}

	return header, body
}

func measure(label string, fn func() error) result {
	for i := 0; i < warmup; i++ {
		must(fn())
	}

	samples := make([]float64, 0, iterations)
	rssDeltas := make([]uint64, 0, iterations)

	for i := 0; i < iterations; i++ {
		before := rssKB()
		started := time.Now()
		must(fn())
		samples = append(samples, time.Since(started).Seconds())

		after := rssKB()
		if after > before {
			rssDeltas = append(rssDeltas, after-before)
		} else {
			rssDeltas = append(rssDeltas, 0)
		}
	}

	mean := 0.0
	minimum := samples[0]
	for _, sample := range samples {
		mean += sample
		if sample < minimum {
			minimum = sample
		}
	}
	mean /= float64(len(samples))

	var variance float64
	for _, sample := range samples {
		diff := sample - mean
		variance += diff * diff
	}
	variance /= float64(len(samples))

	var rssMax uint64
	for _, delta := range rssDeltas {
		if delta > rssMax {
			rssMax = delta
		}
	}

	return result{
		Label:      label,
		Real:       mean,
		RealMin:    minimum,
		RealStddev: math.Sqrt(variance),
		RSSDeltaKB: rssMax,
		Iterations: iterations,
	}
}

func rssKB() uint64 {
	file, err := os.Open("/proc/self/status")
	if err != nil {
		return 0
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		if strings.HasPrefix(line, "VmRSS:") {
			fields := strings.Fields(line)
			if len(fields) >= 2 {
				if value, err := strconv.ParseUint(fields[1], 10, 64); err == nil {
					return value
				}
			}
		}
	}
	return 0
}

func writeWithExcelize(path string, header []any, body [][]any) error {
	book := excelize.NewFile()
	defer func() { _ = book.Close() }()

	sheet := book.GetSheetName(book.GetActiveSheetIndex())
	if err := book.SetSheetName(sheet, "Bench"); err != nil {
		return err
	}

	rows := make([][]any, 0, len(body)+1)
	rows = append(rows, header)
	rows = append(rows, body...)

	for rowIndex, values := range rows {
		cell, err := excelize.CoordinatesToCellName(1, rowIndex+1)
		if err != nil {
			return err
		}
		if err := book.SetSheetRow("Bench", cell, &values); err != nil {
			return err
		}
	}

	return book.SaveAs(path)
}

func readWithExcelize(path string) (int, error) {
	book, err := excelize.OpenFile(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = book.Close() }()

	rows, err := book.Rows("Bench")
	if err != nil {
		return 0, err
	}
	defer func() { _ = rows.Close() }()

	count := 0
	for rows.Next() {
		columns, err := rows.Columns()
		if err != nil {
			return 0, err
		}
		count += len(columns)
	}
	return count, rows.Error()
}

func must(err error) {
	if err != nil {
		fail(err)
	}
}

func fail(err error) {
	fmt.Fprintln(os.Stderr, err)
	os.Exit(1)
}
