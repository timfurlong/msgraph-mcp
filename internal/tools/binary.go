package tools

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

var extByContentType = map[string]string{
	"image/png":  ".png",
	"image/jpeg": ".jpg",
	"image/gif":  ".gif",
	"image/webp": ".webp",
}

// isImage reports a content type that clients can show as an image block.
// The match is exact, as in Python: parameters break it.
func isImage(ct string) bool {
	_, ok := extByContentType[ct]
	return ok
}

func extFor(ct string) string {
	if ext, ok := extByContentType[ct]; ok {
		return ext
	}
	return ".bin"
}

// sniffImage identifies PNG, JPEG, GIF and WebP from magic bytes, since
// Graph's $value endpoints don't say what they return.
func sniffImage(b []byte) string {
	switch {
	case bytes.HasPrefix(b, []byte("\x89PNG\r\n\x1a\n")):
		return "image/png"
	case bytes.HasPrefix(b, []byte{0xff, 0xd8, 0xff}):
		return "image/jpeg"
	case bytes.HasPrefix(b, []byte("GIF87a")), bytes.HasPrefix(b, []byte("GIF89a")):
		return "image/gif"
	case len(b) >= 12 && bytes.Equal(b[:4], []byte("RIFF")) && bytes.Equal(b[8:12], []byte("WEBP")):
		return "image/webp"
	}
	return ""
}

// imageResult is a metadata text block plus a native image block.
func imageResult(meta any, data []byte, mime string) *mcpResult {
	text, err := jsonText(meta)
	if err != nil {
		text = "{}"
	}
	return &mcp.CallToolResult{Content: []mcp.Content{
		&mcp.TextContent{Text: text},
		&mcp.ImageContent{Data: data, MIMEType: mime},
	}}
}

type mcpResult = mcp.CallToolResult

// writeBytes writes data to savePath (~ expanded) and returns the path it
// wrote. If savePath is an existing directory the file goes inside it under
// defaultName, reduced to its base name so it can't point elsewhere. Parent
// directories are created as needed.
func writeBytes(savePath string, data []byte, defaultName string) (string, error) {
	target := savePath
	if target == "~" || strings.HasPrefix(target, "~/") {
		home, err := os.UserHomeDir()
		if err != nil {
			return "", err
		}
		target = filepath.Join(home, strings.TrimPrefix(target, "~"))
	}
	if fi, err := os.Stat(target); err == nil && fi.IsDir() {
		name := filepath.Base(filepath.Clean("/" + filepath.ToSlash(defaultName)))
		if name == "/" || name == "." || name == "" {
			name = "download.bin"
		}
		target = filepath.Join(target, name)
	}
	if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
		return "", err
	}
	if err := os.WriteFile(target, data, 0o644); err != nil {
		return "", err
	}
	return target, nil
}
