// Package trim reshapes Graph JSON into the small, snake_case objects the
// tools return. Every function takes Graph's object as received and, with
// includeRaw, returns that object unchanged under "raw".
package trim

import "encoding/json"

// M is a JSON object.
type M = map[string]any

func withRaw(t, raw M, includeRaw bool) M {
	if includeRaw {
		t["raw"] = raw
	}
	return t
}

// obj returns v as an object, or an empty one.
func obj(v any) M {
	if m, ok := v.(M); ok {
		return m
	}
	return M{}
}

func list(v any) []any {
	if l, ok := v.([]any); ok {
		return l
	}
	return nil
}

func str(v any) string {
	s, _ := v.(string)
	return s
}

// truthy follows Python's bool(): nil, false, "", 0, and empty containers
// are false.
func truthy(v any) bool {
	switch x := v.(type) {
	case nil:
		return false
	case bool:
		return x
	case string:
		return x != ""
	case json.Number:
		f, err := x.Float64()
		return err != nil || f != 0
	case float64:
		return x != 0
	case int:
		return x != 0
	case M:
		return len(x) > 0
	case []any:
		return len(x) > 0
	}
	return true
}

// or returns v unless it is falsy, then def.
func or(v any, def string) any {
	if truthy(v) {
		return v
	}
	return def
}

// email reads a recipient {"emailAddress": {"name", "address"}}.
func email(v any) any {
	if !truthy(v) {
		return nil
	}
	addr := obj(obj(v)["emailAddress"])
	return M{"name": addr["name"], "address": addr["address"]}
}

func emails(v any) []any {
	out := []any{}
	for _, r := range list(v) {
		if e := email(r); e != nil {
			out = append(out, e)
		}
	}
	return out
}

// strList returns a copy of a list, or an empty list.
func strList(v any) []any {
	out := []any{}
	return append(out, list(v)...)
}
