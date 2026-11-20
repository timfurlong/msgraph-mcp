package tools

import (
	"context"
	"encoding/base64"
	"net/url"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

// sortClause leads every list_messages filter. Graph refuses a $filter with
// $orderby unless the filter names the sort property first (400
// InefficientFilter); this always-true clause satisfies that, so filtered
// results stay newest first across pages.
const sortClause = "receivedDateTime ge 1900-01-01T00:00:00Z"

type listMessagesIn struct {
	FolderID   string  `json:"folder_id"`
	Mailbox    *string `json:"mailbox"`
	Limit      int     `json:"limit"`
	PageToken  *string `json:"page_token"`
	IncludeRaw bool    `json:"include_raw"`
	UnreadOnly bool    `json:"unread_only"`
	Filter     *string `json:"filter"`
}

type searchMessagesIn struct {
	Query      string  `json:"query"`
	Mailbox    *string `json:"mailbox"`
	Limit      int     `json:"limit"`
	PageToken  *string `json:"page_token"`
	IncludeRaw bool    `json:"include_raw"`
}

type getMessageIn struct {
	MessageID   string  `json:"message_id"`
	Mailbox     *string `json:"mailbox"`
	IncludeBody bool    `json:"include_body"`
	IncludeRaw  bool    `json:"include_raw"`
}

type listAttachmentsIn struct {
	MessageID  string  `json:"message_id"`
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

type downloadAttachmentIn struct {
	MessageID    string  `json:"message_id"`
	AttachmentID string  `json:"attachment_id"`
	Mailbox      *string `json:"mailbox"`
	SavePath     *string `json:"save_path"`
	IncludeRaw   bool    `json:"include_raw"`
}

func messagePath(mailbox *string, id string) string {
	return graph.Root(mailbox) + "/messages/" + seg(id)
}

func listFilter(unreadOnly bool, filter *string) string {
	var parts []string
	if unreadOnly {
		parts = append(parts, "isRead eq false")
	}
	if filter != nil && *filter != "" {
		parts = append(parts, "("+*filter+")")
	}
	if len(parts) == 0 {
		return ""
	}
	return strings.Join(append([]string{sortClause}, parts...), " and ")
}

// quoteSearch wraps a $search value in double quotes, escaping \ and ".
// Unquoted, a multi-word value comes back unfiltered and a leading digit is
// a syntax error.
func quoteSearch(query string) (string, error) {
	if strings.TrimSpace(query) == "" {
		return "", graph.Invalid("`query` is required")
	}
	if strings.Count(query, `"`)%2 == 1 {
		// Graph answers an unbalanced quote with unfiltered results.
		return "", graph.Invalid("`query` has an unbalanced double quote; close the phrase or drop the \"")
	}
	escaped := strings.ReplaceAll(strings.ReplaceAll(query, `\`, `\\`), `"`, `\"`)
	return `"` + escaped + `"`, nil
}

func messagePage(resp M, includeRaw bool) M {
	out := []any{}
	for _, m := range items(resp) {
		out = append(out, trim.Message(m, false, includeRaw))
	}
	return page(out, resp["@odata.nextLink"])
}

func registerMailRead(s *mcp.Server, g *graph.Client) {
	add(s, "list_messages", func(ctx context.Context, in listMessagesIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		q := url.Values{"$top": {itoa(in.Limit)}, "$orderby": {"receivedDateTime DESC"}}
		if f := listFilter(in.UnreadOnly, in.Filter); f != "" {
			q.Set("$filter", f)
		}
		path := graph.Root(in.Mailbox) + "/mailFolders/" + seg(in.FolderID) + "/messages"
		resp, err := fetchPage(ctx, g, path, q, in.PageToken)
		if err != nil {
			return nil, err
		}
		return messagePage(resp, in.IncludeRaw), nil
	})

	add(s, "search_messages", func(ctx context.Context, in searchMessagesIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		var q url.Values
		// A continuation URL already carries the search, so only a first
		// page needs (and validates) the query.
		if in.PageToken == nil {
			search, err := quoteSearch(in.Query)
			if err != nil {
				return nil, err
			}
			q = url.Values{"$top": {itoa(in.Limit)}, "$search": {search}}
		}
		resp, err := fetchPage(ctx, g, graph.Root(in.Mailbox)+"/messages", q, in.PageToken)
		if err != nil {
			return nil, err
		}
		return messagePage(resp, in.IncludeRaw), nil
	})

	add(s, "get_message", func(ctx context.Context, in getMessageIn) (any, error) {
		msg, err := g.Do(ctx, "GET", messagePath(in.Mailbox, in.MessageID), nil, nil)
		if err != nil {
			return nil, err
		}
		return trim.Message(msg, in.IncludeBody, in.IncludeRaw), nil
	})

	add(s, "list_attachments", func(ctx context.Context, in listAttachmentsIn) (any, error) {
		resp, err := g.Do(ctx, "GET", messagePath(in.Mailbox, in.MessageID)+"/attachments", nil, nil)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, a := range items(resp) {
			out = append(out, trim.AttachmentItem(a, in.IncludeRaw))
		}
		return M{"items": out, "next_page_token": nil}, nil
	})

	add(s, "download_attachment", func(ctx context.Context, in downloadAttachmentIn) (any, error) {
		att, err := g.Do(ctx, "GET", messagePath(in.Mailbox, in.MessageID)+"/attachments/"+seg(in.AttachmentID), nil, nil)
		if err != nil {
			return nil, err
		}
		encoded, hasBytes := att["contentBytes"].(string)
		if hasBytes {
			data, err := base64.StdEncoding.DecodeString(encoded)
			if err != nil {
				return nil, graph.Invalid("attachment content is not valid base64: %v", err)
			}
			name, _ := att["name"].(string)
			ct, _ := att["contentType"].(string)
			if in.SavePath != nil {
				defaultName := name
				if defaultName == "" {
					defaultName = "attachment-" + in.AttachmentID + extFor(ct)
				}
				path, err := writeBytes(*in.SavePath, data, defaultName)
				if err != nil {
					return nil, err
				}
				return M{"path": path, "name": att["name"], "content_type": att["contentType"], "size_bytes": len(data)}, nil
			}
			if isImage(ct) {
				return imageResult(M{"name": att["name"], "content_type": ct, "size_bytes": len(data)}, data, ct), nil
			}
		}
		// Item and reference attachments carry no bytes.
		return trim.AttachmentDownload(att, in.IncludeRaw), nil
	})
}
