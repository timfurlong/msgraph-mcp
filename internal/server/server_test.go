package server

import (
	"strings"
	"testing"
)

func TestInstructionsMentionMailboxAndPaging(t *testing.T) {
	for _, s := range []string{"`mailbox`", "`page_token`", "`include_raw=true`", "download_hosted_content"} {
		if !strings.Contains(Instructions, s) {
			t.Fatalf("instructions missing %s", s)
		}
	}
}
