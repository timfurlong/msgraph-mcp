package auth

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"time"

	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/cache"
)

// FileCache persists MSAL's token cache in one JSON file shared with every
// other copy of the server, including the older Python version, whose MSAL
// reads and writes the same format.
type FileCache struct {
	Path string
	Now  func() time.Time
}

var _ cache.ExportReplace = (*FileCache)(nil)

// Replace loads the file into MSAL's in-memory cache. A missing, empty or
// corrupt file leaves the cache empty, which surfaces as "not authenticated".
func (f *FileCache) Replace(_ context.Context, u cache.Unmarshaler, _ cache.ReplaceHints) error {
	b, err := os.ReadFile(f.Path)
	if err != nil || len(b) == 0 {
		return nil
	}
	pruned, err := pruneRefreshTokens(b)
	if err != nil {
		return nil
	}
	if err := u.Unmarshal(pruned); err != nil {
		return nil
	}
	return nil
}

// Export writes MSAL's cache to the file: mode 0600 inside a 0700 directory,
// through a temp file and a rename so a concurrent reader never sees a
// half-written file.
func (f *FileCache) Export(_ context.Context, m cache.Marshaler, _ cache.ExportHints) error {
	b, err := m.Marshal()
	if err != nil {
		return err
	}
	now := time.Now
	if f.Now != nil {
		now = f.Now
	}
	if b, err = stampRefreshTokens(b, now()); err != nil {
		return err
	}
	dir := filepath.Dir(f.Path)
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return err
	}
	if err := os.Chmod(dir, 0o700); err != nil {
		return err
	}
	tmp := fmt.Sprintf("%s.tmp-%d", f.Path, os.Getpid())
	if err := os.WriteFile(tmp, b, 0o600); err != nil {
		return err
	}
	if err := os.Chmod(tmp, 0o600); err != nil {
		os.Remove(tmp)
		return err
	}
	if err := os.Rename(tmp, f.Path); err != nil {
		os.Remove(tmp)
		return err
	}
	return nil
}

// pruneRefreshTokens keeps only the newest refresh token per account,
// environment and client. MSAL Python keys refresh tokens by scope set, so a
// cache that lived through scope changes holds several, and MSAL Go picks one
// of them arbitrarily, possibly one Entra has already expired for inactivity.
// "Newest" is the largest last_modification_time; an entry without one counts
// as oldest.
func pruneRefreshTokens(b []byte) ([]byte, error) {
	var doc map[string]json.RawMessage
	if err := json.Unmarshal(b, &doc); err != nil {
		return nil, err
	}
	rawRTs, ok := doc["RefreshToken"]
	if !ok {
		return b, nil
	}
	var rts map[string]map[string]any
	if err := json.Unmarshal(rawRTs, &rts); err != nil {
		return nil, err
	}
	type pick struct {
		key  string
		time int64
	}
	best := map[string]pick{}
	for k, rt := range rts {
		group := fmt.Sprint(rt["home_account_id"], "\x00", rt["environment"], "\x00", rt["client_id"])
		t := modTime(rt)
		if cur, ok := best[group]; !ok || t > cur.time || (t == cur.time && k > cur.key) {
			best[group] = pick{k, t}
		}
	}
	kept := make(map[string]map[string]any, len(best))
	for _, p := range best {
		kept[p.key] = rts[p.key]
	}
	var err error
	if doc["RefreshToken"], err = json.Marshal(kept); err != nil {
		return nil, err
	}
	return json.Marshal(doc)
}

// stampRefreshTokens gives every refresh token without a
// last_modification_time the current time. MSAL Go doesn't write the field,
// so this marks its tokens as the newest for pruneRefreshTokens.
func stampRefreshTokens(b []byte, now time.Time) ([]byte, error) {
	var doc map[string]json.RawMessage
	if err := json.Unmarshal(b, &doc); err != nil {
		return nil, err
	}
	rawRTs, ok := doc["RefreshToken"]
	if !ok {
		return b, nil
	}
	var rts map[string]map[string]any
	if err := json.Unmarshal(rawRTs, &rts); err != nil {
		return nil, err
	}
	for _, rt := range rts {
		if _, ok := rt["last_modification_time"]; !ok {
			rt["last_modification_time"] = strconv.FormatInt(now.Unix(), 10)
		}
	}
	var err error
	if doc["RefreshToken"], err = json.Marshal(rts); err != nil {
		return nil, err
	}
	return json.Marshal(doc)
}

func modTime(rt map[string]any) int64 {
	switch v := rt["last_modification_time"].(type) {
	case string:
		n, _ := strconv.ParseInt(v, 10, 64)
		return n
	case float64:
		return int64(v)
	}
	return 0
}
