package tools

import (
	"encoding/base64"
	"os"
	"path/filepath"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func TestIsImageAndExt(t *testing.T) {
	for ct, ext := range map[string]string{"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif", "image/webp": ".webp"} {
		if !isImage(ct) || extFor(ct) != ext {
			t.Fatalf("%s: isImage %v ext %s", ct, isImage(ct), extFor(ct))
		}
	}
	for _, ct := range []string{"", "image/PNG", "image/png; charset=x", "application/pdf", "image/svg+xml"} {
		if isImage(ct) {
			t.Fatalf("%q counted as image", ct)
		}
	}
	if extFor("application/pdf") != ".bin" || extFor("") != ".bin" {
		t.Fatal("non-image ext")
	}
}

func TestSniffImage(t *testing.T) {
	cases := map[string][]byte{
		"image/png":  []byte("\x89PNG\r\n\x1a\nrest"),
		"image/jpeg": {0xff, 0xd8, 0xff, 0xe0},
		"image/gif":  []byte("GIF89a..."),
		"image/webp": []byte("RIFF\x00\x00\x00\x00WEBPVP8 "),
		"":           []byte("%PDF-1.7"),
	}
	cases["image/gif"] = []byte("GIF87a...")
	for want, b := range cases {
		if got := sniffImage(b); got != want {
			t.Fatalf("sniff %q = %q, want %q", b, got, want)
		}
	}
	if sniffImage([]byte("GIF89a")) != "image/gif" {
		t.Fatal("GIF89a")
	}
}

func TestWriteBytesFileAndDirectory(t *testing.T) {
	dir := t.TempDir()
	file := filepath.Join(dir, "nested", "out.bin")
	got, err := writeBytes(file, []byte("abc"), "default.bin")
	if err != nil || got != file {
		t.Fatalf("got %q, %v", got, err)
	}
	if b, _ := os.ReadFile(file); string(b) != "abc" {
		t.Fatalf("content %q", b)
	}
	got, err = writeBytes(dir, []byte("xyz"), "named.txt")
	if err != nil || got != filepath.Join(dir, "named.txt") {
		t.Fatalf("got %q, %v", got, err)
	}
}

func TestWriteBytesDirectoryStripsTraversal(t *testing.T) {
	dir := t.TempDir()
	got, err := writeBytes(dir, []byte("x"), "../../evil.txt")
	if err != nil || got != filepath.Join(dir, "evil.txt") {
		t.Fatalf("got %q, %v", got, err)
	}
}

func TestWriteBytesExpandsTilde(t *testing.T) {
	home := t.TempDir()
	t.Setenv("HOME", home)
	got, err := writeBytes("~/dl/a.txt", []byte("x"), "d")
	if err != nil || got != filepath.Join(home, "dl", "a.txt") {
		t.Fatalf("got %q, %v", got, err)
	}
}

func TestImageResult(t *testing.T) {
	res := imageResult(map[string]any{"name": "p.png", "size_bytes": 3}, []byte{1, 2, 3}, "image/png")
	if len(res.Content) != 2 {
		t.Fatalf("%d blocks", len(res.Content))
	}
	if _, ok := res.Content[0].(*mcp.TextContent); !ok {
		t.Fatal("first block is not text")
	}
	img, ok := res.Content[1].(*mcp.ImageContent)
	if !ok || img.MIMEType != "image/png" || base64.StdEncoding.EncodeToString(img.Data) != "AQID" {
		t.Fatalf("image block %+v", res.Content[1])
	}
}
