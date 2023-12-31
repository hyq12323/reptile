package slogx

import (
	"fmt"
	"io"
	"log"
	"log/slog"
	"os"
	"path/filepath"
	"sync"
	"time"
)

var logger *MyLogger
var writer io.Writer

type MyLogger struct {
	*slog.Logger
	slog.Level
}

type RotateWriter struct {
	baseDir  string
	dateDir  string
	dnFormat string
	fileName string
	fnFormat string
	ext      string
	maxBytes int64

	fd *os.File
	mu sync.Mutex
}

func SetDefault(debug bool) {
	if debug {
		writer = os.Stdout
		NewLogger(writer, slog.LevelDebug)
	} else {
		writer = &RotateWriter{
			baseDir:  "log",
			dateDir:  "",
			dnFormat: "200601",
			fileName: "",
			fnFormat: "02",
			ext:      ".log",
			maxBytes: 10 * 1024 * 1024,
		}
		NewLogger(writer, slog.LevelInfo)
	}
}

func NewLogger(w io.Writer, level slog.Level) {
	opts := &slog.HandlerOptions{
		AddSource: true,
		Level:     level,
		ReplaceAttr: func(groups []string, a slog.Attr) slog.Attr {
			if a.Key == slog.TimeKey {
				a.Value = slog.StringValue(a.Value.Time().Format(time.DateTime))
			}
			if a.Key == slog.SourceKey {
				source := a.Value.Any().(*slog.Source)
				source.File = filepath.Base(source.File)
			}
			return a
		},
	}
	handler := slog.NewTextHandler(w, opts)
	logger = &MyLogger{
		Logger: slog.New(handler),
		Level:  slog.LevelInfo,
	}
	slog.SetDefault(logger.Logger)
	return
}
func LogLogger() *log.Logger {
	return slog.NewLogLogger(logger.Handler(), logger.Level)
}
func Writer() io.Writer {
	return writer
}
func GormLogLevel() int {
	switch logger.Level {
	case slog.LevelDebug:
		return 1
	case slog.LevelInfo:
		return 4
	case slog.LevelWarn:
		return 3
	case slog.LevelError:
		return 2
	}
	return 1
}

func (w *RotateWriter) Write(p []byte) (n int, err error) {
	w.mu.Lock()
	defer w.mu.Unlock()

	w.dateDir = time.Now().Format(w.dnFormat)
	dirPath := filepath.Join(w.baseDir, w.dateDir)
	if _, err := os.Stat(dirPath); os.IsNotExist(err) {
		_ = os.MkdirAll(dirPath, 0666)
		w.openFile()
	} else {
		if fi, err := w.fd.Stat(); err == nil {
			fName := time.Now().Format(w.fnFormat)
			if fName != w.fileName {
				_ = w.fd.Close()
				w.openFile()
			} else if fi.Size() > w.maxBytes {
				_ = w.fd.Close()
				w.rotate()
				w.openFile()
			}
		} else {
			w.openFile()
		}
	}
	return w.fd.Write(p)
}

func (w *RotateWriter) openFile() {
	w.fileName = time.Now().Format(w.fnFormat)
	fPath := filepath.Join(w.baseDir, w.dateDir, w.fileName+w.ext)
	w.fd, _ = os.OpenFile(fPath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
}

func (w *RotateWriter) rotate() {
	oldPath := filepath.Join(w.baseDir, w.dateDir, w.fileName+w.ext)
	newFileName := fmt.Sprintf("%s_%d%s", w.fileName, time.Now().Unix(), w.ext)
	newPath := filepath.Join(w.baseDir, w.dateDir, newFileName)
	_ = os.Rename(oldPath, newPath)
}
