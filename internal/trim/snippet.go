package trim

import (
	"encoding/json"
	"html"
	"regexp"
	"strings"
	"unicode"
)

const snippetLimit = 280

var (
	tagRE    = regexp.MustCompile(`<[^>]+>`)
	hostedRE = regexp.MustCompile(`hostedContents/([^/"'\s]+)/\$value`)
)

// isSpace matches what Python's str.isspace() accepts, so collapsing and
// stripping treat non-breaking and other Unicode spaces the same way.
func isSpace(r rune) bool {
	return unicode.IsSpace(r) || (r >= 0x1c && r <= 0x1f)
}

func collapse(s string) string {
	return strings.Join(strings.FieldsFunc(s, isSpace), " ")
}

// truncate cuts to snippetLimit characters (runes), trims trailing space and
// adds "...".
func truncate(s string) string {
	r := []rune(s)
	if len(r) <= snippetLimit {
		return s
	}
	return strings.TrimRightFunc(string(r[:snippetLimit]), isSpace) + "..."
}

// Snippet is a short plain-text preview of a body: tags removed for HTML,
// entities decoded, whitespace collapsed, cut at 280 characters.
func Snippet(content, contentType string) *string {
	if content == "" {
		return nil
	}
	text := content
	if strings.ToLower(contentType) == "html" {
		text = tagRE.ReplaceAllString(text, " ")
	}
	text = truncate(collapse(html.UnescapeString(text)))
	if text == "" {
		return nil
	}
	return &text
}

// CardText builds a snippet from the "text" values of card attachments, for
// app and bot posts whose body is just an <attachment> tag.
func CardText(attachments []any) *string {
	var texts []string
	for _, a := range attachments {
		am := obj(a)
		content, ok := am["content"].(string)
		if !ok || !strings.Contains(str(am["contentType"]), "card") {
			continue
		}
		found, err := cardTexts(content)
		if err != nil {
			continue
		}
		texts = append(texts, found...)
	}
	var kept []string
	for _, t := range texts {
		if strings.TrimFunc(t, isSpace) != "" {
			kept = append(kept, t)
		}
	}
	joined := truncate(collapse(strings.Join(kept, " ")))
	if joined == "" {
		return nil
	}
	return &joined
}

// cardTexts walks card JSON in document order and collects every string
// value under a "text" key. Non-string "text" values are walked into.
func cardTexts(content string) ([]string, error) {
	dec := json.NewDecoder(strings.NewReader(content))
	var out []string
	var walk func() error
	walk = func() error {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch tok {
		case json.Delim('{'):
			for dec.More() {
				keyTok, err := dec.Token()
				if err != nil {
					return err
				}
				if key, _ := keyTok.(string); key == "text" {
					var v json.RawMessage
					if err := dec.Decode(&v); err != nil {
						return err
					}
					var s string
					if json.Unmarshal(v, &s) == nil {
						out = append(out, s)
						continue
					}
					sub, err := cardTexts(string(v))
					if err != nil {
						return err
					}
					out = append(out, sub...)
					continue
				}
				if err := walk(); err != nil {
					return err
				}
			}
			_, err = dec.Token() // }
			return err
		case json.Delim('['):
			for dec.More() {
				if err := walk(); err != nil {
					return err
				}
			}
			_, err = dec.Token() // ]
			return err
		}
		return nil
	}
	if err := walk(); err != nil {
		return nil, err
	}
	return out, nil
}

// HostedRefs lists the hosted content ids referenced in a body, first seen
// first, without duplicates.
func HostedRefs(content string) []any {
	out := []any{}
	seen := map[string]bool{}
	for _, m := range hostedRE.FindAllStringSubmatch(content, -1) {
		if !seen[m[1]] {
			seen[m[1]] = true
			out = append(out, M{"hosted_content_id": m[1]})
		}
	}
	return out
}
