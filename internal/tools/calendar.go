package tools

import (
	"context"
	"net/url"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

type listCalendarsIn struct {
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

type listEventsIn struct {
	CalendarID    string  `json:"calendar_id"`
	StartDatetime *string `json:"start_datetime"`
	EndDatetime   *string `json:"end_datetime"`
	Mailbox       *string `json:"mailbox"`
	Limit         int     `json:"limit"`
	PageToken     *string `json:"page_token"`
	IncludeRaw    bool    `json:"include_raw"`
}

type getEventIn struct {
	EventID     string  `json:"event_id"`
	Mailbox     *string `json:"mailbox"`
	IncludeBody bool    `json:"include_body"`
	IncludeRaw  bool    `json:"include_raw"`
}

type reminderFields struct {
	ShowAs                     *string `json:"show_as"`
	IsReminderOn               *bool   `json:"is_reminder_on"`
	ReminderMinutesBeforeStart *int    `json:"reminder_minutes_before_start"`
}

type createEventIn struct {
	reminderFields
	Subject         string   `json:"subject"`
	StartDatetime   string   `json:"start_datetime"`
	EndDatetime     string   `json:"end_datetime"`
	TimeZone        string   `json:"time_zone"`
	Body            *string  `json:"body"`
	BodyType        string   `json:"body_type"`
	Location        *string  `json:"location"`
	Attendees       []string `json:"attendees"`
	IsOnlineMeeting bool     `json:"is_online_meeting"`
	IsAllDay        bool     `json:"is_all_day"`
	CalendarID      *string  `json:"calendar_id"`
	Mailbox         *string  `json:"mailbox"`
	IncludeRaw      bool     `json:"include_raw"`
}

type updateEventIn struct {
	reminderFields
	EventID         string  `json:"event_id"`
	Subject         *string `json:"subject"`
	StartDatetime   *string `json:"start_datetime"`
	EndDatetime     *string `json:"end_datetime"`
	TimeZone        *string `json:"time_zone"`
	Body            *string `json:"body"`
	BodyType        string  `json:"body_type"`
	Location        *string `json:"location"`
	IsOnlineMeeting *bool   `json:"is_online_meeting"`
	IsAllDay        *bool   `json:"is_all_day"`
	Mailbox         *string `json:"mailbox"`
	IncludeRaw      bool    `json:"include_raw"`
}

type eventIDIn struct {
	EventID string  `json:"event_id"`
	Mailbox *string `json:"mailbox"`
}

type cancelEventIn struct {
	EventID string  `json:"event_id"`
	Comment *string `json:"comment"`
	Mailbox *string `json:"mailbox"`
}

type respondIn struct {
	EventID      string  `json:"event_id"`
	Response     string  `json:"response"`
	Comment      *string `json:"comment"`
	SendResponse bool    `json:"send_response"`
	Mailbox      *string `json:"mailbox"`
}

type findMeetingTimesIn struct {
	Attendees       []string `json:"attendees"`
	DurationMinutes int      `json:"duration_minutes"`
	StartWindow     string   `json:"start_window"`
	EndWindow       string   `json:"end_window"`
	Mailbox         *string  `json:"mailbox"`
	MaxCandidates   int      `json:"max_candidates"`
	IncludeRaw      bool     `json:"include_raw"`
}

var responseStatus = map[string]string{"accept": "accepted", "tentativelyAccept": "tentativelyAccepted", "decline": "declined"}

func eventPath(mailbox *string, id string) string {
	return graph.Root(mailbox) + "/events/" + seg(id)
}

func dateTimeTZ(dt, tz string) M { return M{"dateTime": dt, "timeZone": tz} }

func itemBody(text, bodyType string) M {
	ct := "text"
	if bodyType == "html" {
		ct = "html"
	}
	return M{"contentType": ct, "content": text}
}

// apply sets showAs and the reminder fields on an event body.
func (r reminderFields) apply(evt M) error {
	if r.ShowAs != nil {
		if err := oneOf("show_as", *r.ShowAs, "busy", "free", "oof", "tentative", "workingElsewhere"); err != nil {
			return err
		}
		evt["showAs"] = *r.ShowAs
	}
	switch {
	case r.ReminderMinutesBeforeStart != nil:
		if *r.ReminderMinutesBeforeStart < 0 {
			return graph.Invalid("reminder_minutes_before_start must be zero or greater")
		}
		if r.IsReminderOn != nil && !*r.IsReminderOn {
			return graph.Invalid("reminder_minutes_before_start cannot be set when is_reminder_on is False")
		}
		evt["reminderMinutesBeforeStart"] = *r.ReminderMinutesBeforeStart
		evt["isReminderOn"] = true
	case r.IsReminderOn != nil:
		evt["isReminderOn"] = *r.IsReminderOn
	}
	return nil
}

func slot(v any) any {
	s, ok := v.(M)
	if !ok {
		return nil
	}
	return M{"date_time": s["dateTime"], "time_zone": s["timeZone"]}
}

func registerCalendar(s *mcp.Server, g *graph.Client) {
	add(s, "list_calendars", func(ctx context.Context, in listCalendarsIn) (any, error) {
		resp, err := g.Do(ctx, "GET", graph.Root(in.Mailbox)+"/calendars", nil, nil)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, c := range items(resp) {
			out = append(out, trim.Calendar(c, in.IncludeRaw))
		}
		return M{"items": out, "next_page_token": nil}, nil
	})

	add(s, "list_events", func(ctx context.Context, in listEventsIn) (any, error) {
		if in.StartDatetime == nil || *in.StartDatetime == "" || in.EndDatetime == nil || *in.EndDatetime == "" {
			return nil, graph.Invalid("start_datetime and end_datetime are required")
		}
		if err := graph.ValidateLimit(in.Limit, 100); err != nil {
			return nil, err
		}
		q := url.Values{
			"startDateTime": {*in.StartDatetime},
			"endDateTime":   {*in.EndDatetime},
			"$top":          {itoa(in.Limit)},
			"$orderby":      {"start/dateTime"},
		}
		path := graph.Root(in.Mailbox) + "/calendars/" + seg(in.CalendarID) + "/calendarView"
		resp, err := fetchPage(ctx, g, path, q, in.PageToken)
		if err != nil {
			return nil, err
		}
		out := []any{}
		for _, e := range items(resp) {
			out = append(out, trim.Event(e, false, in.IncludeRaw))
		}
		return page(out, resp["@odata.nextLink"]), nil
	})

	add(s, "get_event", func(ctx context.Context, in getEventIn) (any, error) {
		evt, err := g.Do(ctx, "GET", eventPath(in.Mailbox, in.EventID), nil, nil)
		if err != nil {
			return nil, err
		}
		return trim.Event(evt, in.IncludeBody, in.IncludeRaw), nil
	})

	add(s, "create_event", func(ctx context.Context, in createEventIn) (any, error) {
		evt := M{
			"subject":         in.Subject,
			"start":           dateTimeTZ(in.StartDatetime, in.TimeZone),
			"end":             dateTimeTZ(in.EndDatetime, in.TimeZone),
			"isOnlineMeeting": in.IsOnlineMeeting,
			"isAllDay":        in.IsAllDay,
		}
		if in.Body != nil {
			evt["body"] = itemBody(*in.Body, in.BodyType)
		}
		if in.Location != nil {
			evt["location"] = M{"displayName": *in.Location}
		}
		if len(in.Attendees) > 0 {
			atts := []any{}
			for _, a := range in.Attendees {
				atts = append(atts, M{"emailAddress": M{"address": a}, "type": "required"})
			}
			evt["attendees"] = atts
		}
		if err := in.apply(evt); err != nil {
			return nil, err
		}
		path := graph.Root(in.Mailbox) + "/events"
		if in.CalendarID != nil {
			path = graph.Root(in.Mailbox) + "/calendars/" + seg(*in.CalendarID) + "/events"
		}
		created, err := g.Do(ctx, "POST", path, nil, evt)
		if err != nil {
			return nil, err
		}
		return trim.Event(created, false, in.IncludeRaw), nil
	})

	add(s, "update_event", func(ctx context.Context, in updateEventIn) (any, error) {
		hasTZ := in.TimeZone != nil && *in.TimeZone != ""
		if ((in.StartDatetime != nil && *in.StartDatetime != "") || (in.EndDatetime != nil && *in.EndDatetime != "")) && !hasTZ {
			return nil, graph.Invalid("time_zone is required when updating start or end")
		}
		tz := "UTC"
		if hasTZ {
			tz = *in.TimeZone
		}
		patch := M{}
		if in.Subject != nil {
			patch["subject"] = *in.Subject
		}
		if in.StartDatetime != nil {
			patch["start"] = dateTimeTZ(*in.StartDatetime, tz)
		}
		if in.EndDatetime != nil {
			patch["end"] = dateTimeTZ(*in.EndDatetime, tz)
		}
		if in.Body != nil {
			patch["body"] = itemBody(*in.Body, in.BodyType)
		}
		if in.Location != nil {
			patch["location"] = M{"displayName": *in.Location}
		}
		if in.IsOnlineMeeting != nil {
			patch["isOnlineMeeting"] = *in.IsOnlineMeeting
		}
		if in.IsAllDay != nil {
			patch["isAllDay"] = *in.IsAllDay
		}
		if err := in.apply(patch); err != nil {
			return nil, err
		}
		updated, err := g.Do(ctx, "PATCH", eventPath(in.Mailbox, in.EventID), nil, patch)
		if err != nil {
			return nil, err
		}
		return trim.Event(updated, false, in.IncludeRaw), nil
	})

	add(s, "delete_event", func(ctx context.Context, in eventIDIn) (any, error) {
		if _, err := g.Do(ctx, "DELETE", eventPath(in.Mailbox, in.EventID), nil, nil); err != nil {
			return nil, err
		}
		return M{"status": "deleted"}, nil
	})

	add(s, "cancel_event", func(ctx context.Context, in cancelEventIn) (any, error) {
		body := M{}
		if in.Comment != nil {
			body["comment"] = *in.Comment
		}
		if _, err := g.Do(ctx, "POST", eventPath(in.Mailbox, in.EventID)+"/cancel", nil, body); err != nil {
			return nil, err
		}
		return M{"status": "cancelled"}, nil
	})

	add(s, "respond_to_event", func(ctx context.Context, in respondIn) (any, error) {
		status, ok := responseStatus[in.Response]
		if !ok {
			return nil, oneOf("response", in.Response, "accept", "decline", "tentativelyAccept")
		}
		// No comment means no comment key at all.
		body := M{"sendResponse": in.SendResponse}
		if in.Comment != nil {
			body["comment"] = *in.Comment
		}
		if _, err := g.Do(ctx, "POST", eventPath(in.Mailbox, in.EventID)+"/"+in.Response, nil, body); err != nil {
			return nil, err
		}
		return M{"status": status}, nil
	})

	add(s, "find_meeting_times", func(ctx context.Context, in findMeetingTimesIn) (any, error) {
		if len(in.Attendees) == 0 {
			return nil, graph.Invalid("`attendees` must be a non-empty list")
		}
		if in.DurationMinutes <= 0 {
			return nil, graph.Invalid("duration_minutes must be a positive integer")
		}
		atts := []any{}
		for _, a := range in.Attendees {
			atts = append(atts, M{"emailAddress": M{"address": a}})
		}
		body := M{
			"attendees":       atts,
			"meetingDuration": "PT" + itoa(in.DurationMinutes) + "M",
			"maxCandidates":   in.MaxCandidates,
			"timeConstraint": M{"timeSlots": []any{M{
				"start": dateTimeTZ(in.StartWindow, "UTC"),
				"end":   dateTimeTZ(in.EndWindow, "UTC"),
			}}},
		}
		result, err := g.Do(ctx, "POST", graph.Root(in.Mailbox)+"/findMeetingTimes", nil, body)
		if err != nil {
			return nil, err
		}
		suggestions := []any{}
		for _, raw := range trimList(result["meetingTimeSuggestions"]) {
			ts, _ := raw["meetingTimeSlot"].(M)
			suggestions = append(suggestions, M{
				"start":      slot(ts["start"]),
				"end":        slot(ts["end"]),
				"confidence": raw["confidence"],
				"order_hint": raw["order"],
			})
		}
		out := M{"suggestions": suggestions, "empty_reason": result["emptySuggestionsReason"]}
		if in.IncludeRaw {
			out["raw"] = result
		}
		return out, nil
	})
}

func trimList(v any) []M {
	raw, _ := v.([]any)
	out := make([]M, 0, len(raw))
	for _, x := range raw {
		if m, ok := x.(M); ok {
			out = append(out, m)
		}
	}
	return out
}
