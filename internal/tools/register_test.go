package tools

import (
	"context"
	"encoding/json"
	"reflect"
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/tools/toolspec"
)

func TestAllSnapshotToolsRegistered(t *testing.T) {
	cs, _ := newSession(t, nil)
	res, err := cs.ListTools(context.Background(), nil)
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]int{}
	for i, tool := range res.Tools {
		got[tool.Name] = i
	}
	specs := toolspec.All()
	if len(res.Tools) != len(specs) {
		t.Fatalf("%d tools listed, want %d", len(res.Tools), len(specs))
	}
	for _, spec := range specs {
		i, ok := got[spec.Name]
		if !ok {
			t.Fatalf("%s not registered", spec.Name)
		}
		tool := res.Tools[i]
		if tool.Description != spec.Description {
			t.Fatalf("%s: description differs", spec.Name)
		}
		var want, have any
		_ = json.Unmarshal(spec.InputSchema, &want)
		b, _ := json.Marshal(tool.InputSchema)
		_ = json.Unmarshal(b, &have)
		if !reflect.DeepEqual(want, have) {
			t.Fatalf("%s: input schema differs\n got %s\nwant %s", spec.Name, b, spec.InputSchema)
		}
	}
}
