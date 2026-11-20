package tools

import (
	"context"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

type ruleFields struct {
	SenderContains        []string `json:"sender_contains"`
	SubjectContains       []string `json:"subject_contains"`
	BodyContains          []string `json:"body_contains"`
	BodyOrSubjectContains []string `json:"body_or_subject_contains"`
	FromAddresses         []string `json:"from_addresses"`
	HasAttachments        *bool    `json:"has_attachments"`
	MoveToFolder          *string  `json:"move_to_folder"`
	MarkAsRead            *bool    `json:"mark_as_read"`
	Delete                *bool    `json:"delete"`
	StopProcessingRules   *bool    `json:"stop_processing_rules"`
}

type listRulesIn struct {
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

type ruleIDIn struct {
	RuleID     string  `json:"rule_id"`
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

type createRuleIn struct {
	ruleFields
	DisplayName string  `json:"display_name"`
	Sequence    *int    `json:"sequence"`
	IsEnabled   bool    `json:"is_enabled"`
	Mailbox     *string `json:"mailbox"`
	IncludeRaw  bool    `json:"include_raw"`
}

type updateRuleIn struct {
	ruleFields
	RuleID      string  `json:"rule_id"`
	DisplayName *string `json:"display_name"`
	IsEnabled   *bool   `json:"is_enabled"`
	Sequence    *int    `json:"sequence"`
	Mailbox     *string `json:"mailbox"`
	IncludeRaw  bool    `json:"include_raw"`
}

func rulesPath(mailbox *string) string {
	return graph.Root(mailbox) + "/mailFolders/inbox/messageRules"
}

// predicates builds the conditions object with only the fields that have
// values (an explicit null breaks Graph's rule endpoints).
func (f ruleFields) predicates() (M, error) {
	p := M{}
	if len(f.SenderContains) > 0 {
		p["senderContains"] = f.SenderContains
	}
	if len(f.SubjectContains) > 0 {
		p["subjectContains"] = f.SubjectContains
	}
	if len(f.BodyContains) > 0 {
		p["bodyContains"] = f.BodyContains
	}
	if len(f.BodyOrSubjectContains) > 0 {
		p["bodyOrSubjectContains"] = f.BodyOrSubjectContains
	}
	if len(f.FromAddresses) > 0 {
		for _, a := range f.FromAddresses {
			if a == "" || !strings.Contains(a, "@") {
				return nil, graph.Invalid("Invalid email address: %s", pyRepr(a))
			}
		}
		p["fromAddresses"] = recipients(f.FromAddresses)
	}
	if f.HasAttachments != nil {
		p["hasAttachments"] = *f.HasAttachments
	}
	if len(p) == 0 {
		return nil, nil
	}
	return p, nil
}

func (f ruleFields) actions() M {
	a := M{}
	if f.MoveToFolder != nil {
		a["moveToFolder"] = *f.MoveToFolder
	}
	if f.MarkAsRead != nil {
		a["markAsRead"] = *f.MarkAsRead
	}
	if f.Delete != nil {
		a["delete"] = *f.Delete
	}
	if f.StopProcessingRules != nil {
		a["stopProcessingRules"] = *f.StopProcessingRules
	}
	if len(a) == 0 {
		return nil
	}
	return a
}

func registerMailRules(s *mcp.Server, g *graph.Client) {
	add(s, "list_rules", func(ctx context.Context, in listRulesIn) (any, error) {
		resp, err := g.Do(ctx, "GET", rulesPath(in.Mailbox), nil, nil)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, r := range items(resp) {
			out = append(out, trim.Rule(r, in.IncludeRaw))
		}
		return M{"items": out, "next_page_token": nil}, nil
	})

	add(s, "get_rule", func(ctx context.Context, in ruleIDIn) (any, error) {
		if in.RuleID == "" {
			return nil, graph.Invalid("`rule_id` is required")
		}
		rule, err := g.Do(ctx, "GET", rulesPath(in.Mailbox)+"/"+seg(in.RuleID), nil, nil)
		if err != nil {
			return nil, err
		}
		return trim.Rule(rule, in.IncludeRaw), nil
	})

	add(s, "create_rule", func(ctx context.Context, in createRuleIn) (any, error) {
		if strings.TrimSpace(in.DisplayName) == "" {
			return nil, graph.Invalid("`display_name` is required")
		}
		conditions, err := in.predicates()
		if err != nil {
			return nil, err
		}
		if conditions == nil {
			return nil, graph.Invalid("At least one condition is required (sender_contains, subject_contains, body_contains, body_or_subject_contains, from_addresses, or has_attachments)")
		}
		actions := in.actions()
		if actions == nil {
			return nil, graph.Invalid("At least one action is required (move_to_folder, mark_as_read, delete, or stop_processing_rules)")
		}
		// Graph rejects sequence 0, so an unset sequence is 1.
		sequence := 1
		if in.Sequence != nil {
			sequence = *in.Sequence
		}
		body := M{"displayName": in.DisplayName, "isEnabled": in.IsEnabled, "conditions": conditions, "actions": actions, "sequence": sequence}
		created, err := g.Do(ctx, "POST", rulesPath(in.Mailbox), nil, body)
		if err != nil {
			return nil, err
		}
		return trim.Rule(created, in.IncludeRaw), nil
	})

	add(s, "update_rule", func(ctx context.Context, in updateRuleIn) (any, error) {
		if in.RuleID == "" {
			return nil, graph.Invalid("`rule_id` is required")
		}
		conditions, err := in.predicates()
		if err != nil {
			return nil, err
		}
		actions := in.actions()
		if in.DisplayName == nil && in.IsEnabled == nil && in.Sequence == nil && conditions == nil && actions == nil {
			return nil, graph.Invalid("At least one field must be updated (display_name, is_enabled, sequence, a condition, or an action)")
		}
		patch := M{}
		if in.DisplayName != nil {
			if strings.TrimSpace(*in.DisplayName) == "" {
				return nil, graph.Invalid("`display_name` cannot be empty")
			}
			patch["displayName"] = *in.DisplayName
		}
		if in.IsEnabled != nil {
			patch["isEnabled"] = *in.IsEnabled
		}
		if in.Sequence != nil {
			patch["sequence"] = *in.Sequence
		}
		if conditions != nil {
			patch["conditions"] = conditions
		}
		if actions != nil {
			patch["actions"] = actions
		}
		updated, err := g.Do(ctx, "PATCH", rulesPath(in.Mailbox)+"/"+seg(in.RuleID), nil, patch)
		if err != nil {
			return nil, err
		}
		return trim.Rule(updated, in.IncludeRaw), nil
	})

	add(s, "delete_rule", func(ctx context.Context, in ruleIDIn) (any, error) {
		if in.RuleID == "" {
			return nil, graph.Invalid("`rule_id` is required")
		}
		if _, err := g.Do(ctx, "DELETE", rulesPath(in.Mailbox)+"/"+seg(in.RuleID), nil, nil); err != nil {
			return nil, err
		}
		return M{"deleted": true, "id": in.RuleID}, nil
	})
}
