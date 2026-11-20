package graph

import "fmt"

// APIError is a non-success response from Graph, or a failure to reach it
// (Status 0). Its text is what callers see.
type APIError struct {
	Status  int
	Code    string
	Message string
}

func (e *APIError) Error() string {
	return fmt.Sprintf("Graph API %d: %s — %s", e.Status, e.Code, e.Message)
}

// ValidationError is a bad tool argument caught before any Graph call.
type ValidationError struct{ Msg string }

func (e *ValidationError) Error() string { return e.Msg }

// Invalid returns a ValidationError with a formatted message.
func Invalid(format string, args ...any) error {
	return &ValidationError{Msg: fmt.Sprintf(format, args...)}
}
