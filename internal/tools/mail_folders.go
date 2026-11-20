package tools

import (
	"context"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

type listFoldersIn struct {
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

type moveMessageIn struct {
	MessageID   string  `json:"message_id"`
	Destination string  `json:"destination"`
	Mailbox     *string `json:"mailbox"`
	IncludeRaw  bool    `json:"include_raw"`
}

type createFolderIn struct {
	DisplayName    string  `json:"display_name"`
	ParentFolderID *string `json:"parent_folder_id"`
	Mailbox        *string `json:"mailbox"`
	IncludeRaw     bool    `json:"include_raw"`
}

type updateFolderIn struct {
	FolderID       string  `json:"folder_id"`
	DisplayName    *string `json:"display_name"`
	ParentFolderID *string `json:"parent_folder_id"`
	Mailbox        *string `json:"mailbox"`
	IncludeRaw     bool    `json:"include_raw"`
}

type deleteFolderIn struct {
	FolderID string  `json:"folder_id"`
	Mailbox  *string `json:"mailbox"`
}

func folderPath(mailbox *string, id string) string {
	return graph.Root(mailbox) + "/mailFolders/" + seg(id)
}

// moveMessage moves a message; Graph resolves well-known folder names
// (archive, deleteditems, ...) itself.
func moveMessage(ctx context.Context, g *graph.Client, in moveMessageIn) (any, error) {
	if in.Destination == "" {
		return nil, graph.Invalid("`destination` is required")
	}
	moved, err := g.Do(ctx, "POST", messagePath(in.Mailbox, in.MessageID)+"/move", nil, M{"destinationId": in.Destination})
	if err != nil {
		return nil, err
	}
	return trim.Message(moved, false, in.IncludeRaw), nil
}

func number(v any) int {
	n, _ := strconvAtoi(v)
	return n
}

func registerMailFolders(s *mcp.Server, g *graph.Client) {
	add(s, "list_folders", func(ctx context.Context, in listFoldersIn) (any, error) {
		resp, err := g.Do(ctx, "GET", graph.Root(in.Mailbox)+"/mailFolders", nil, nil)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, f := range items(resp) {
			out = append(out, trim.Folder(f, in.IncludeRaw))
		}
		return M{"items": out, "next_page_token": nil}, nil
	})

	add(s, "move_message", func(ctx context.Context, in moveMessageIn) (any, error) {
		return moveMessage(ctx, g, in)
	})

	add(s, "create_folder", func(ctx context.Context, in createFolderIn) (any, error) {
		if strings.TrimSpace(in.DisplayName) == "" {
			return nil, graph.Invalid("`display_name` is required")
		}
		path := graph.Root(in.Mailbox) + "/mailFolders"
		if in.ParentFolderID != nil && *in.ParentFolderID != "" {
			path = folderPath(in.Mailbox, *in.ParentFolderID) + "/childFolders"
		}
		created, err := g.Do(ctx, "POST", path, nil, M{"displayName": in.DisplayName})
		if err != nil {
			return nil, err
		}
		return trim.Folder(created, in.IncludeRaw), nil
	})

	add(s, "update_folder", func(ctx context.Context, in updateFolderIn) (any, error) {
		if in.FolderID == "" {
			return nil, graph.Invalid("`folder_id` is required")
		}
		if in.DisplayName == nil && in.ParentFolderID == nil {
			return nil, graph.Invalid("At least one of `display_name` or `parent_folder_id` is required")
		}
		if in.DisplayName != nil && strings.TrimSpace(*in.DisplayName) == "" {
			return nil, graph.Invalid("`display_name` cannot be empty")
		}
		path := folderPath(in.Mailbox, in.FolderID)
		var result M
		var err error
		if in.DisplayName != nil {
			if result, err = g.Do(ctx, "PATCH", path, nil, M{"displayName": *in.DisplayName}); err != nil {
				return nil, err
			}
		}
		if in.ParentFolderID != nil {
			if result, err = g.Do(ctx, "POST", path+"/move", nil, M{"destinationId": *in.ParentFolderID}); err != nil {
				return nil, err
			}
		}
		return trim.Folder(result, in.IncludeRaw), nil
	})

	add(s, "delete_folder", func(ctx context.Context, in deleteFolderIn) (any, error) {
		if in.FolderID == "" {
			return nil, graph.Invalid("`folder_id` is required")
		}
		path := folderPath(in.Mailbox, in.FolderID)
		folder, err := g.Do(ctx, "GET", path, nil, nil)
		if err != nil {
			return nil, err
		}
		total, children := number(folder["totalItemCount"]), number(folder["childFolderCount"])
		if total > 0 || children > 0 {
			return nil, graph.Invalid("Refusing to delete non-empty folder %s: contains %d message(s) and %d child folder(s). Empty it first or move the contents elsewhere.", pyRepr(in.FolderID), total, children)
		}
		if _, err := g.Do(ctx, "DELETE", path, nil, nil); err != nil {
			return nil, err
		}
		return M{"deleted": true, "id": in.FolderID}, nil
	})
}
