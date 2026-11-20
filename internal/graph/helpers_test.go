package graph_test

import (
	"errors"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

func asValidation(err error, v **graph.ValidationError) bool { return errors.As(err, v) }
