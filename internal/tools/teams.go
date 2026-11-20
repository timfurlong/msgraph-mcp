package tools

import (
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"net/url"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

// Graph caps chat, channel message and reply pages at 50.
const chatMessageMax = 50

type pagedIn struct {
	Limit      int     `json:"limit"`
	PageToken  *string `json:"page_token"`
	IncludeRaw bool    `json:"include_raw"`
}

type chatMessagesIn struct {
	pagedIn
	ChatID      string `json:"chat_id"`
	IncludeBody bool   `json:"include_body"`
}

type channelsIn struct {
	pagedIn
	TeamID string `json:"team_id"`
}

type channelMessagesIn struct {
	pagedIn
	TeamID      string `json:"team_id"`
	ChannelID   string `json:"channel_id"`
	IncludeBody bool   `json:"include_body"`
}

type repliesIn struct {
	channelMessagesIn
	MessageID string `json:"message_id"`
}

type hostedContentIn struct {
	MessageID       string  `json:"message_id"`
	HostedContentID string  `json:"hosted_content_id"`
	ChatID          *string `json:"chat_id"`
	TeamID          *string `json:"team_id"`
	ChannelID       *string `json:"channel_id"`
	SavePath        *string `json:"save_path"`
	IncludeRaw      bool    `json:"include_raw"`
}

func chatMessagePage(resp M, includeBody, includeRaw bool) M {
	out := []any{}
	for _, m := range items(resp) {
		out = append(out, trim.ChatMessage(m, includeBody, includeRaw))
	}
	return page(out, resp["@odata.nextLink"])
}

func top(limit int) url.Values { return url.Values{"$top": {itoa(limit)}} }

func channelPath(teamID, channelID string) string {
	return "/teams/" + seg(teamID) + "/channels/" + seg(channelID)
}

func registerTeams(s *mcp.Server, g *graph.Client) {
	add(s, "list_chats", func(ctx context.Context, in pagedIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		// lastMessagePreview gives the real last activity; lastUpdatedDateTime
		// only moves on renames and membership changes.
		q := url.Values{"$top": {itoa(in.Limit)}, "$expand": {"members,lastMessagePreview"}}
		resp, err := fetchPage(ctx, g, "/me/chats", q, in.PageToken)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, c := range items(resp) {
			out = append(out, trim.Chat(c, in.IncludeRaw))
		}
		return page(out, resp["@odata.nextLink"]), nil
	})

	add(s, "list_chat_messages", func(ctx context.Context, in chatMessagesIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, chatMessageMax); err != nil {
			return nil, err
		}
		resp, err := fetchPage(ctx, g, "/chats/"+seg(in.ChatID)+"/messages", top(in.Limit), in.PageToken)
		if err != nil {
			return nil, err
		}
		return chatMessagePage(resp, in.IncludeBody, in.IncludeRaw), nil
	})

	add(s, "list_joined_teams", func(ctx context.Context, in pagedIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		// /me/joinedTeams rejects $top, so fetch unpaged and cap here.
		resp, err := fetchPage(ctx, g, "/me/joinedTeams", nil, in.PageToken)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, t := range capped(items(resp), in.Limit) {
			out = append(out, trim.Team(t, in.IncludeRaw))
		}
		return page(out, resp["@odata.nextLink"]), nil
	})

	add(s, "list_channels", func(ctx context.Context, in channelsIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		// /teams/{id}/channels rejects $top, so fetch unpaged and cap here.
		resp, err := fetchPage(ctx, g, "/teams/"+seg(in.TeamID)+"/channels", nil, in.PageToken)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, c := range capped(items(resp), in.Limit) {
			out = append(out, trim.Channel(c, in.IncludeRaw))
		}
		return page(out, resp["@odata.nextLink"]), nil
	})

	add(s, "list_channel_messages", func(ctx context.Context, in channelMessagesIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, chatMessageMax); err != nil {
			return nil, err
		}
		resp, err := fetchPage(ctx, g, channelPath(in.TeamID, in.ChannelID)+"/messages", top(in.Limit), in.PageToken)
		if err != nil {
			return nil, err
		}
		return chatMessagePage(resp, in.IncludeBody, in.IncludeRaw), nil
	})

	add(s, "list_message_replies", func(ctx context.Context, in repliesIn) (any, error) {
		if err := graph.ValidateLimit(in.Limit, chatMessageMax); err != nil {
			return nil, err
		}
		path := channelPath(in.TeamID, in.ChannelID) + "/messages/" + seg(in.MessageID) + "/replies"
		resp, err := fetchPage(ctx, g, path, top(in.Limit), in.PageToken)
		if err != nil {
			return nil, err
		}
		return chatMessagePage(resp, in.IncludeBody, in.IncludeRaw), nil
	})

	add(s, "download_hosted_content", func(ctx context.Context, in hostedContentIn) (any, error) {
		hasChat, hasTeam, hasChannel := in.ChatID != nil, in.TeamID != nil, in.ChannelID != nil
		if hasChat && (hasTeam || hasChannel) {
			return nil, graph.Invalid("Provide either chat_id or (team_id and channel_id), not both.")
		}
		if !hasChat && !(hasTeam && hasChannel) {
			return nil, graph.Invalid("Provide a message location: chat_id, or both team_id and channel_id.")
		}
		var msgPath string
		if hasChat {
			msgPath = "/chats/" + seg(*in.ChatID) + "/messages/" + seg(in.MessageID)
		} else {
			msgPath = channelPath(*in.TeamID, *in.ChannelID) + "/messages/" + seg(in.MessageID)
		}
		data, err := g.GetBytes(ctx, msgPath+"/hostedContents/"+seg(in.HostedContentID)+"/$value")
		if err != nil {
			return nil, err
		}
		// $value doesn't say what it returns, so sniff the bytes.
		ct := sniffImage(data)
		var ctVal any
		if ct != "" {
			ctVal = ct
		}
		if in.SavePath != nil {
			// The hosted content id is opaque base64, unsafe as a file name.
			sum := sha256.Sum256([]byte(in.HostedContentID))
			name := "hosted-content-" + in.MessageID + "-" + hex.EncodeToString(sum[:])[:8] + extFor(ct)
			path, err := writeBytes(*in.SavePath, data, name)
			if err != nil {
				return nil, err
			}
			return M{"path": path, "content_type": ctVal, "size_bytes": len(data)}, nil
		}
		if isImage(ct) {
			return imageResult(M{"content_type": ct, "size_bytes": len(data)}, data, ct), nil
		}
		raw := M{"contentType": ctVal, "size": len(data), "contentBytes": base64.StdEncoding.EncodeToString(data)}
		return trim.HostedDownload(raw, in.IncludeRaw), nil
	})
}

func capped(xs []M, n int) []M {
	if len(xs) > n {
		return xs[:n]
	}
	return xs
}
