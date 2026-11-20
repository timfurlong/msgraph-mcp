// Package toolspec holds the name, description and input schema of every
// tool, as a snapshot of the Python server's tools/list (v0.5.4) with two
// deliberate edits. Registering from it keeps the Go tools' schemas identical
// to what clients already know.
package toolspec

import (
	_ "embed"
	"encoding/json"
	"sync"
)

// Tool is one tool's definition as clients see it in tools/list.
type Tool struct {
	Name        string          `json:"name"`
	Description string          `json:"description"`
	InputSchema json.RawMessage `json:"inputSchema"`
}

//go:embed tools.json
var raw []byte

var (
	once   sync.Once
	tools  []Tool
	byName map[string]Tool
)

func load() {
	if err := json.Unmarshal(raw, &tools); err != nil {
		panic("toolspec: bad tools.json: " + err.Error())
	}
	byName = make(map[string]Tool, len(tools))
	for _, t := range tools {
		byName[t.Name] = t
	}
}

// All returns every tool in snapshot order.
func All() []Tool {
	once.Do(load)
	return tools
}

// Get returns the named tool.
func Get(name string) (Tool, bool) {
	once.Do(load)
	t, ok := byName[name]
	return t, ok
}
