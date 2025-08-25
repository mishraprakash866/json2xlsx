package main

import (
	"bufio"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log"
	"mime"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Row is a generic JSON object
type Row = map[string]any

func main() {
	mux := http.NewServeMux()
	// Upload a JSON file (multipart/form-data key "file") OR raw application/json body.
	mux.HandleFunc("/upload-json-to-xlsx", uploadJSONToXLSX)

	s := &http.Server{
		Addr:              ":8080",
		Handler:           logMiddleware(mux),
		ReadHeaderTimeout: 10 * time.Second,
		ReadTimeout:       0, // large uploads: don't cap; rely on infra/proxy
		WriteTimeout:      0, // large downloads
		IdleTimeout:       120 * time.Second,
	}
	log.Printf("Listening on http://localhost%s …", s.Addr)
	log.Fatal(s.ListenAndServe())
}

func logMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		st := time.Now()
		next.ServeHTTP(w, r)
		log.Printf("%s %s %s", r.Method, r.URL.Path, time.Since(st))
	})
}

// ---------- Handler

func uploadJSONToXLSX(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Use POST. Content-Type: multipart/form-data (key=file) or application/json", http.StatusMethodNotAllowed)
		return
	}

	// Hard cap uploads to 1 GB
	r.Body = http.MaxBytesReader(w, r.Body, 1<<30)

	// Save incoming body to a temp file (so we can do two streaming passes)
	tmpPath, filename, err := persistUploadToTemp(r)
	if err != nil {
		http.Error(w, "Upload error: "+err.Error(), http.StatusBadRequest)
		return
	}
	defer os.Remove(tmpPath)

	// PASS 1: discover headers per sheet
	headersBySheet, order, err := discoverHeaders(tmpPath)
	if err != nil {
		http.Error(w, "Invalid JSON file: "+err.Error(), http.StatusBadRequest)
		return
	}
	if len(headersBySheet) == 0 {
		http.Error(w, "JSON contained no rows", http.StatusBadRequest)
		return
	}

	// PASS 2: write XLSX via stream writer
	f := excelize.NewFile()

	// Remove default sheet only if we will create our own
	defaultSheet := f.GetSheetName(f.GetActiveSheetIndex())
	if len(order) > 0 && !(len(order) == 1 && order[0] == defaultSheet) {
		// create sheets first; we’ll delete default if unused
	}

	for i, sheet := range order {
		if i == 0 && sheet == defaultSheet {
			// reuse default sheet name
		} else {
			f.NewSheet(sheet)
		}
		hdrs := headersBySheet[sheet]
		// stream
		sw, err := f.NewStreamWriter(sheet)
		if err != nil {
			http.Error(w, "Failed to init sheet writer: "+err.Error(), http.StatusInternalServerError)
			return
		}
		// header row
		headerRow := make([]any, len(hdrs))
		for i := range hdrs {
			headerRow[i] = hdrs[i]
		}
		if err := sw.SetRow("A1", headerRow); err != nil {
			http.Error(w, "Failed to write header row: "+err.Error(), http.StatusInternalServerError)
			return
		}
		if err := streamRows(tmpPath, sheet, hdrs, sw); err != nil {
			http.Error(w, "Failed to write rows: "+err.Error(), http.StatusBadRequest)
			return
		}
		if err := sw.Flush(); err != nil {
			http.Error(w, "Failed to finalize sheet: "+err.Error(), http.StatusInternalServerError)
			return
		}
	}

	// Set active sheet to first in order
	if len(order) > 0 {
		if idx, err := f.GetSheetIndex(order[0]); err == nil {
			f.SetActiveSheet(idx)
		}
	}

	// Name
	outName := strings.TrimSuffix(filename, filepath.Ext(filename))
	if outName == "" || outName == "." || outName == "/" {
		outName = "output"
	}
	outName += ".xlsx"

	// Stream XLSX to client
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename=%q`, outName))
	if err := f.Write(w); err != nil {
		// client may disconnect; just log
		log.Printf("write error: %v", err)
	}
}

// ---------- Upload persistence

func persistUploadToTemp(r *http.Request) (tmpPath string, filename string, err error) {
	ct := r.Header.Get("Content-Type")
	mediaType, params, _ := mime.ParseMediaType(ct)

	// Create temp file
	tmpFile, err := os.CreateTemp("", "json2xlsx-*.json")
	if err != nil {
		return "", "", err
	}
	defer func() {
		if err != nil {
			tmpFile.Close()
			os.Remove(tmpFile.Name())
		}
	}()

	switch {
	case strings.HasPrefix(mediaType, "multipart/"):
		if err := r.ParseMultipartForm(64 << 20); err != nil { // 64MB form mem before spilling to disk
			return "", "", fmt.Errorf("parse multipart: %w", err)
		}
		file, hdr, err := r.FormFile("file")
		if err != nil {
			return "", "", errors.New("multipart: expected form-data key 'file'")
		}
		defer file.Close()
		if _, err := io.Copy(tmpFile, file); err != nil {
			return "", "", fmt.Errorf("copy upload: %w", err)
		}
		filename = hdr.Filename

	case mediaType == "application/json" || mediaType == "text/json" || mediaType == "":
		// Treat raw body as JSON
		if _, err := io.Copy(tmpFile, r.Body); err != nil {
			return "", "", fmt.Errorf("read body: %w", err)
		}
		filename = "data.json"

	default:
		_ = params // unused
		return "", "", fmt.Errorf("unsupported Content-Type %q", mediaType)
	}

	if err := tmpFile.Close(); err != nil {
		return "", "", err
	}
	return tmpFile.Name(), filename, nil
}

// ---------- Pass 1: discover headers

// discoverHeaders returns: map[sheetName][]headers and the sheet order to produce.
func discoverHeaders(path string) (map[string][]string, []string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	// limited := io.LimitReader(f, 600<<20) // max 10MB per JSON object
	// dec := json.NewDecoder(limited)
	dec := json.NewDecoder(bufio.NewReader(f))
	tok, err := dec.Token()
	if err != nil {
		return nil, nil, errors.New("unable to read JSON (is the file empty?)")
	}

	headersBySheet := make(map[string]map[string]struct{}) // set
	order := []string{}

	switch d := tok.(type) {
	case json.Delim:
		if d == '[' {
			// Single-sheet: array of objects → "Sheet1"
			sheet := "Sheet1"
			order = append(order, sheet)
			headersBySheet[sheet] = make(map[string]struct{})
			for dec.More() {
				var m Row
				if err := dec.Decode(&m); err != nil {
					return nil, nil, fmt.Errorf("invalid JSON row in array: %w", err)
				}
				for k := range m {
					headersBySheet[sheet][k] = struct{}{}
				}
			}
			// consume closing ']'
			if err := consumeArrayEnd(dec); err != nil {
				return nil, nil, err
			}
		} else if d == '{' {
			// Multi-sheet: object where each value is an array of objects
			for dec.More() {
				kTok, err := dec.Token()
				if err != nil {
					return nil, nil, errors.New("invalid JSON object: missing key")
				}
				sheet, ok := kTok.(string)
				if !ok {
					return nil, nil, errors.New("invalid JSON: non-string key at top level")
				}
				// Next token must be '[' (array)
				if err := expectArrayStart(dec); err != nil {
					return nil, nil, fmt.Errorf("sheet %q is not a JSON array: %w", sheet, err)
				}
				if _, exists := headersBySheet[sheet]; !exists {
					headersBySheet[sheet] = make(map[string]struct{})
					order = append(order, sheet)
				}
				for dec.More() {
					var m Row
					if err := dec.Decode(&m); err != nil {
						return nil, nil, fmt.Errorf("invalid row in sheet %q: %w", sheet, err)
					}
					for k := range m {
						headersBySheet[sheet][k] = struct{}{}
					}
				}
				if err := consumeArrayEnd(dec); err != nil {
					return nil, nil, fmt.Errorf("sheet %q: %w", sheet, err)
				}
			}
			// consume closing '}'
			if err := expectObjectEnd(dec); err != nil {
				return nil, nil, err
			}
		} else {
			return nil, nil, errors.New("top-level JSON must be an array or object")
		}
	default:
		return nil, nil, errors.New("top-level JSON must be an array or object")
	}

	// Convert header sets → sorted slices
	final := make(map[string][]string, len(headersBySheet))
	for sheet, set := range headersBySheet {
		h := make([]string, 0, len(set))
		for k := range set {
			h = append(h, k)
		}
		sort.Strings(h)
		final[sheet] = h
	}
	return final, order, nil
}

func expectArrayStart(dec *json.Decoder) error {
	t, err := dec.Token()
	if err != nil {
		return err
	}
	d, ok := t.(json.Delim)
	if !ok || d != '[' {
		return errors.New("expected '['")
	}
	return nil
}
func consumeArrayEnd(dec *json.Decoder) error {
	t, err := dec.Token()
	if err != nil {
		return err
	}
	d, ok := t.(json.Delim)
	if !ok || d != ']' {
		return errors.New("expected ']'")
	}
	return nil
}
func expectObjectEnd(dec *json.Decoder) error {
	t, err := dec.Token()
	if err != nil {
		return err
	}
	d, ok := t.(json.Delim)
	if !ok || d != '}' {
		return errors.New("expected '}'")
	}
	return nil
}

// ---------- Pass 2: stream rows into XLSX

func streamRows(path, targetSheet string, headers []string, sw *excelize.StreamWriter) error {
	f, err := os.Open(path)
	if err != nil {
		return err
	}
	defer f.Close()

	dec := json.NewDecoder(bufio.NewReader(f))
	tok, err := dec.Token()
	if err != nil {
		return errors.New("unable to read JSON")
	}

	rowIdx := 2 // 1-based; row 1 is headers

	switch d := tok.(type) {
	case json.Delim:
		if d == '[' {
			// Single sheet
			if targetSheet != "Sheet1" {
				// This pass is for a different sheet; consume and return without writing
				// (We call this function once per sheet; only write when matching)
				// Just drain array:
				for dec.More() {
					var skip any
					if err := dec.Decode(&skip); err != nil {
						return err
					}
				}
				return consumeArrayEnd(dec)
			}
			for dec.More() {
				var m Row
				if err := dec.Decode(&m); err != nil {
					return fmt.Errorf("invalid JSON row: %w", err)
				}
				row := buildRow(headers, m)
				cell, _ := excelize.CoordinatesToCellName(1, rowIdx)
				if err := sw.SetRow(cell, row); err != nil {
					return err
				}
				rowIdx++
			}
			return consumeArrayEnd(dec)
		} else if d == '{' {
			// Multi-sheet
			for dec.More() {
				kTok, err := dec.Token()
				if err != nil {
					return err
				}
				sheet, ok := kTok.(string)
				if !ok {
					return errors.New("non-string key at top level")
				}
				if err := expectArrayStart(dec); err != nil {
					return err
				}
				if sheet != targetSheet {
					// drain array
					for dec.More() {
						var skip any
						if err := dec.Decode(&skip); err != nil {
							return err
						}
					}
					if err := consumeArrayEnd(dec); err != nil {
						return err
					}
					continue
				}
				// write rows for the matching sheet
				for dec.More() {
					var m Row
					if err := dec.Decode(&m); err != nil {
						return err
					}
					row := buildRow(headers, m)
					cell, _ := excelize.CoordinatesToCellName(1, rowIdx)
					if err := sw.SetRow(cell, row); err != nil {
						return err
					}
					rowIdx++
				}
				if err := consumeArrayEnd(dec); err != nil {
					return err
				}
			}
			return expectObjectEnd(dec)
		}
	}
	return errors.New("top-level JSON must be an array or object")
}

func buildRow(headers []string, m Row) []any {
	row := make([]any, len(headers))
	for i, h := range headers {
		row[i] = m[h]
	}
	return row
}
