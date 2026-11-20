package tools

import (
	"context"
	"encoding/base64"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

// maxInlineAttachmentBytes caps attachments sent inline with a message;
// larger files need an upload session, which isn't supported.
const maxInlineAttachmentBytes = 3 * 1024 * 1024

type composeIn struct {
	To          []string         `json:"to"`
	Subject     string           `json:"subject"`
	Body        string           `json:"body"`
	BodyType    string           `json:"body_type"`
	Cc          []string         `json:"cc"`
	Bcc         []string         `json:"bcc"`
	Attachments []map[string]any `json:"attachments"`
	Mailbox     *string          `json:"mailbox"`
}

type sendMessageIn struct {
	composeIn
	SaveToSentItems bool `json:"save_to_sent_items"`
}

type createDraftIn struct {
	composeIn
	IncludeRaw bool `json:"include_raw"`
}

type replyIn struct {
	MessageID string   `json:"message_id"`
	Comment   *string  `json:"comment"`
	ExtraTo   []string `json:"extra_to"`
	Mailbox   *string  `json:"mailbox"`
}

type replyAllIn struct {
	MessageID string  `json:"message_id"`
	Comment   *string `json:"comment"`
	Mailbox   *string `json:"mailbox"`
}

type forwardIn struct {
	MessageID string   `json:"message_id"`
	To        []string `json:"to"`
	Comment   *string  `json:"comment"`
	Mailbox   *string  `json:"mailbox"`
}

type updateMessageIn struct {
	MessageID      string   `json:"message_id"`
	IsRead         *bool    `json:"is_read"`
	Flag           *string  `json:"flag"`
	Importance     *string  `json:"importance"`
	Categories     []string `json:"categories"`
	ParentFolderID *string  `json:"parent_folder_id"`
	Mailbox        *string  `json:"mailbox"`
	IncludeRaw     bool     `json:"include_raw"`
}

type messageIDIn struct {
	MessageID string  `json:"message_id"`
	Mailbox   *string `json:"mailbox"`
}

func recipients(addrs []string) []any {
	out := []any{}
	for _, a := range addrs {
		out = append(out, M{"emailAddress": M{"address": a}})
	}
	return out
}

var errNoRecipients = graph.Invalid("`to` must be a non-empty list of email addresses")

func buildAttachments(specs []map[string]any) ([]any, error) {
	var out []any
	total := 0
	for _, spec := range specs {
		encoded, ok := spec["content_b64"].(string)
		if !ok {
			return nil, graph.Invalid("Invalid attachment content_b64: missing")
		}
		raw, err := base64.StdEncoding.DecodeString(encoded)
		if err != nil {
			return nil, graph.Invalid("Invalid attachment content_b64: %v", err)
		}
		total += len(raw)
		if total > maxInlineAttachmentBytes {
			return nil, graph.Invalid("Attachment payload exceeds 3 MB raw limit (got %d bytes). Use a chunked upload (not supported in v1) for larger files.", total)
		}
		name, _ := spec["name"].(string)
		if name == "" {
			name = "attachment"
		}
		ct, _ := spec["content_type"].(string)
		if ct == "" {
			ct = "application/octet-stream"
		}
		out = append(out, M{
			"@odata.type":  "#microsoft.graph.fileAttachment",
			"name":         name,
			"contentType":  ct,
			"contentBytes": base64.StdEncoding.EncodeToString(raw),
		})
	}
	return out, nil
}

func buildMessage(in composeIn) (M, error) {
	bodyType := "text"
	if in.BodyType == "html" {
		bodyType = "html"
	}
	msg := M{
		"subject":       in.Subject,
		"body":          M{"contentType": bodyType, "content": in.Body},
		"toRecipients":  recipients(in.To),
		"ccRecipients":  recipients(in.Cc),
		"bccRecipients": recipients(in.Bcc),
	}
	atts, err := buildAttachments(in.Attachments)
	if err != nil {
		return nil, err
	}
	if len(atts) > 0 {
		msg["attachments"] = atts
	}
	return msg, nil
}

func oneOf(field, value string, allowed ...string) error {
	for _, a := range allowed {
		if value == a {
			return nil
		}
	}
	list := ""
	for i, a := range allowed {
		if i > 0 {
			list += ", "
		}
		list += "'" + a + "'"
	}
	return graph.Invalid("%s must be one of [%s]", field, list)
}

var sent = M{"status": "sent"}

func registerMailWrite(s *mcp.Server, g *graph.Client) {
	add(s, "send_message", func(ctx context.Context, in sendMessageIn) (any, error) {
		if len(in.To) == 0 {
			return nil, errNoRecipients
		}
		msg, err := buildMessage(in.composeIn)
		if err != nil {
			return nil, err
		}
		body := M{"message": msg, "saveToSentItems": in.SaveToSentItems}
		if _, err := g.Do(ctx, "POST", graph.Root(in.Mailbox)+"/sendMail", nil, body); err != nil {
			return nil, err
		}
		return sent, nil
	})

	add(s, "create_draft", func(ctx context.Context, in createDraftIn) (any, error) {
		msg, err := buildMessage(in.composeIn)
		if err != nil {
			return nil, err
		}
		created, err := g.Do(ctx, "POST", graph.Root(in.Mailbox)+"/messages", nil, msg)
		if err != nil {
			return nil, err
		}
		return trim.Message(created, false, in.IncludeRaw), nil
	})

	add(s, "reply_message", func(ctx context.Context, in replyIn) (any, error) {
		body := M{}
		if in.Comment != nil {
			body["comment"] = *in.Comment
		}
		if len(in.ExtraTo) > 0 {
			body["message"] = M{"toRecipients": recipients(in.ExtraTo)}
		}
		if _, err := g.Do(ctx, "POST", messagePath(in.Mailbox, in.MessageID)+"/reply", nil, body); err != nil {
			return nil, err
		}
		return sent, nil
	})

	add(s, "reply_all_message", func(ctx context.Context, in replyAllIn) (any, error) {
		body := M{}
		if in.Comment != nil {
			body["comment"] = *in.Comment
		}
		if _, err := g.Do(ctx, "POST", messagePath(in.Mailbox, in.MessageID)+"/replyAll", nil, body); err != nil {
			return nil, err
		}
		return sent, nil
	})

	add(s, "forward_message", func(ctx context.Context, in forwardIn) (any, error) {
		if len(in.To) == 0 {
			return nil, errNoRecipients
		}
		body := M{"toRecipients": recipients(in.To)}
		if in.Comment != nil {
			body["comment"] = *in.Comment
		}
		if _, err := g.Do(ctx, "POST", messagePath(in.Mailbox, in.MessageID)+"/forward", nil, body); err != nil {
			return nil, err
		}
		return sent, nil
	})

	add(s, "update_message", func(ctx context.Context, in updateMessageIn) (any, error) {
		return updateMessage(ctx, g, in)
	})

	add(s, "delete_message", func(ctx context.Context, in messageIDIn) (any, error) {
		if _, err := g.Do(ctx, "DELETE", messagePath(in.Mailbox, in.MessageID), nil, nil); err != nil {
			return nil, err
		}
		return M{"status": "deleted"}, nil
	})
}

func updateMessage(ctx context.Context, g *graph.Client, in updateMessageIn) (any, error) {
	patch := M{}
	if in.IsRead != nil {
		patch["isRead"] = *in.IsRead
	}
	if in.Flag != nil {
		if err := oneOf("flag", *in.Flag, "notFlagged", "flagged", "complete"); err != nil {
			return nil, err
		}
		patch["flag"] = M{"flagStatus": *in.Flag}
	}
	if in.Importance != nil {
		if err := oneOf("importance", *in.Importance, "low", "normal", "high"); err != nil {
			return nil, err
		}
		patch["importance"] = *in.Importance
	}
	if in.Categories != nil {
		patch["categories"] = in.Categories
	}
	if in.ParentFolderID != nil {
		patch["parentFolderId"] = *in.ParentFolderID
	}
	updated, err := g.Do(ctx, "PATCH", messagePath(in.Mailbox, in.MessageID), nil, patch)
	if err != nil {
		return nil, err
	}
	return trim.Message(updated, false, in.IncludeRaw), nil
}
