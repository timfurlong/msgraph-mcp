package tools

import (
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func eventJSON() M {
	return M{"id": "e1", "subject": "Standup", "start": M{"dateTime": "2026-05-19T15:00:00", "timeZone": "UTC"}, "end": M{"dateTime": "2026-05-19T15:30:00", "timeZone": "UTC"}, "isReminderOn": true}
}

func TestListCalendars(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "c1", "name": "Calendar", "owner": M{"name": "A", "address": "a@example.com"}, "canEdit": true, "isDefaultCalendar": true}}})
	})
	got := callJSON(t, cs, "list_calendars", M{"mailbox": "shared@example.com"})
	if f.Last().RawPath != "/users/shared%40example.com/calendars" || got["items"].([]any)[0].(M)["name"] != "Calendar" || got["next_page_token"] != nil {
		t.Fatalf("%s %v", f.Last().RawPath, got)
	}
}

func TestListEventsCalendarView(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{eventJSON()}, "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendars/c1/calendarView?$skip=25"})
	})
	got := callJSON(t, cs, "list_events", M{"calendar_id": "c1", "start_datetime": "2026-05-19T00:00:00Z", "end_datetime": "2026-05-20T00:00:00Z"})
	r := f.Last()
	if r.Path != "/me/calendars/c1/calendarView" || r.Query.Get("startDateTime") != "2026-05-19T00:00:00Z" || r.Query.Get("endDateTime") != "2026-05-20T00:00:00Z" ||
		r.Query.Get("$top") != "25" || r.Query.Get("$orderby") != "start/dateTime" {
		t.Fatalf("%s %v", r.Path, r.Query)
	}
	if got["items"].([]any)[0].(M)["subject"] != "Standup" || got["next_page_token"] == nil {
		t.Fatalf("got %v", got)
	}
}

func TestListEventsRequiresBothDatetimes(t *testing.T) {
	cs, _ := newSession(t, nil)
	if msg := callErr(t, cs, "list_events", M{"calendar_id": "c1", "start_datetime": "2026-05-19T00:00:00Z", "end_datetime": ""}); msg != "start_datetime and end_datetime are required" {
		t.Fatalf("msg %q", msg)
	}
}

func TestGetEvent(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		e := eventJSON()
		e["body"] = M{"contentType": "html", "content": "<p>x</p>"}
		return jsonOK(e)
	})
	got := callJSON(t, cs, "get_event", M{"event_id": "e1", "include_body": true})
	if f.Last().Path != "/me/events/e1" || got["body"] != "<p>x</p>" {
		t.Fatalf("%s %v", f.Last().Path, got)
	}
}

func created(graphtest.Request) graphtest.Response {
	return graphtest.Response{Status: 201, JSON: eventJSON()}
}

func TestCreateEventMinimal(t *testing.T) {
	cs, f := newSession(t, created)
	got := callJSON(t, cs, "create_event", M{"subject": "Standup", "start_datetime": "2026-05-19T15:00:00", "end_datetime": "2026-05-19T15:30:00"})
	if got["id"] != "e1" {
		t.Fatalf("got %v", got)
	}
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/events" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "event", r.Body, M{
		"subject":         "Standup",
		"start":           M{"dateTime": "2026-05-19T15:00:00", "timeZone": "UTC"},
		"end":             M{"dateTime": "2026-05-19T15:30:00", "timeZone": "UTC"},
		"isOnlineMeeting": false,
		"isAllDay":        false,
	})
}

func TestCreateEventFull(t *testing.T) {
	cs, f := newSession(t, created)
	callJSON(t, cs, "create_event", M{"subject": "S", "start_datetime": "s", "end_datetime": "e", "time_zone": "America/Denver",
		"body": "<b>x</b>", "body_type": "html", "location": "Room", "attendees": []any{"a@example.com"}, "is_online_meeting": true,
		"show_as": "tentative", "reminder_minutes_before_start": 10, "calendar_id": "c9"})
	r := f.Last()
	if r.Path != "/me/calendars/c9/events" {
		t.Fatalf("path %s", r.Path)
	}
	eq(t, "event", r.Body, M{
		"subject": "S", "start": M{"dateTime": "s", "timeZone": "America/Denver"}, "end": M{"dateTime": "e", "timeZone": "America/Denver"},
		"body": M{"contentType": "html", "content": "<b>x</b>"}, "location": M{"displayName": "Room"},
		"attendees":       []any{M{"emailAddress": M{"address": "a@example.com"}, "type": "required"}},
		"isOnlineMeeting": true, "isAllDay": false, "showAs": "tentative", "isReminderOn": true, "reminderMinutesBeforeStart": 10,
	})
}

func TestCreateEventReminderOff(t *testing.T) {
	cs, f := newSession(t, created)
	callJSON(t, cs, "create_event", M{"subject": "S", "start_datetime": "s", "end_datetime": "e", "is_reminder_on": false})
	if f.Last().Body["isReminderOn"] != false {
		t.Fatalf("body %v", f.Last().Body)
	}
	if _, has := f.Last().Body["reminderMinutesBeforeStart"]; has {
		t.Fatal("minutes set")
	}
}

func TestEventValidation(t *testing.T) {
	cs, f := newSession(t, created)
	base := func(extra M) M {
		m := M{"subject": "S", "start_datetime": "s", "end_datetime": "e"}
		for k, v := range extra {
			m[k] = v
		}
		return m
	}
	for want, args := range map[string]M{
		"reminder_minutes_before_start must be zero or greater":                    base(M{"reminder_minutes_before_start": -1}),
		"reminder_minutes_before_start cannot be set when is_reminder_on is False": base(M{"reminder_minutes_before_start": 5, "is_reminder_on": false}),
	} {
		if msg := callErr(t, cs, "create_event", args); msg != want {
			t.Fatalf("msg %q, want %q", msg, want)
		}
	}
	if msg := callErr(t, cs, "update_event", M{"event_id": "e1", "start_datetime": "s"}); msg != "time_zone is required when updating start or end" {
		t.Fatalf("msg %q", msg)
	}
	if len(f.Requests) != 0 {
		t.Fatal("called Graph")
	}
}

func TestUpdateEvent(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(eventJSON()) })
	callJSON(t, cs, "update_event", M{"event_id": "e1", "subject": "New", "show_as": "oof", "reminder_minutes_before_start": 0})
	r := f.Last()
	if r.Method != "PATCH" || r.Path != "/me/events/e1" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "patch", r.Body, M{"subject": "New", "showAs": "oof", "isReminderOn": true, "reminderMinutesBeforeStart": 0})
	callJSON(t, cs, "update_event", M{"event_id": "e1", "start_datetime": "s", "end_datetime": "e", "time_zone": "UTC", "is_reminder_on": false, "location": "L", "body": "b"})
	eq(t, "patch 2", f.Last().Body, M{"start": M{"dateTime": "s", "timeZone": "UTC"}, "end": M{"dateTime": "e", "timeZone": "UTC"}, "isReminderOn": false, "location": M{"displayName": "L"}, "body": M{"contentType": "text", "content": "b"}})
}

func TestDeleteAndCancelEvent(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 202} })
	eq(t, "delete", callJSON(t, cs, "delete_event", M{"event_id": "e1"}), M{"status": "deleted"})
	if f.Last().Method != "DELETE" || f.Last().Path != "/me/events/e1" {
		t.Fatalf("%s %s", f.Last().Method, f.Last().Path)
	}
	eq(t, "cancel", callJSON(t, cs, "cancel_event", M{"event_id": "e1", "comment": "Sorry"}), M{"status": "cancelled"})
	if r := f.Last(); r.Path != "/me/events/e1/cancel" || r.Body["comment"] != "Sorry" {
		t.Fatalf("%s %v", r.Path, r.Body)
	}
	callJSON(t, cs, "cancel_event", M{"event_id": "e1"})
	if string(f.Last().RawBody) != "{}" {
		t.Fatalf("body %s", f.Last().RawBody)
	}
}

func TestRespondToEvent(t *testing.T) {
	cases := []struct {
		response, path, status string
		args                   M
		body                   string
	}{
		{"accept", "/me/events/e1/accept", "accepted", nil, `{"sendResponse":true}`},
		{"tentativelyAccept", "/me/events/e1/tentativelyAccept", "tentativelyAccepted", M{"comment": "See you there"}, `{"comment":"See you there","sendResponse":true}`},
		{"decline", "/me/events/e1/decline", "declined", M{"send_response": false}, `{"sendResponse":false}`},
	}
	for _, c := range cases {
		cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 202} })
		args := M{"event_id": "e1", "response": c.response}
		for k, v := range c.args {
			args[k] = v
		}
		eq(t, c.response, callJSON(t, cs, "respond_to_event", args), M{"status": c.status})
		if r := f.Last(); r.Path != c.path || string(r.RawBody) != c.body {
			t.Fatalf("%s: %s %s", c.response, r.Path, r.RawBody)
		}
	}
}

func TestFindMeetingTimes(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"@odata.context": "ctx", "emptySuggestionsReason": "", "meetingTimeSuggestions": []any{
			M{"confidence": 100.0, "order": 1, "meetingTimeSlot": M{"start": M{"dateTime": "2026-05-20T16:00:00.0000000", "timeZone": "UTC"}, "end": M{"dateTime": "2026-05-20T16:30:00.0000000", "timeZone": "UTC"}}},
		}})
	})
	got := callJSON(t, cs, "find_meeting_times", M{"attendees": []any{"a@example.com"}, "duration_minutes": 90, "start_window": "2026-05-20T09:00:00", "end_window": "2026-05-20T17:00:00", "include_raw": true})
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/findMeetingTimes" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "body", r.Body, M{
		"attendees":       []any{M{"emailAddress": M{"address": "a@example.com"}}},
		"meetingDuration": "PT90M",
		"maxCandidates":   20,
		"timeConstraint":  M{"timeSlots": []any{M{"start": M{"dateTime": "2026-05-20T09:00:00", "timeZone": "UTC"}, "end": M{"dateTime": "2026-05-20T17:00:00", "timeZone": "UTC"}}}},
	})
	if got["empty_reason"] != "" || got["raw"].(M)["@odata.context"] != "ctx" {
		t.Fatalf("got %v", got)
	}
	eq(t, "suggestions", got["suggestions"], []any{M{
		"start":      M{"date_time": "2026-05-20T16:00:00.0000000", "time_zone": "UTC"},
		"end":        M{"date_time": "2026-05-20T16:30:00.0000000", "time_zone": "UTC"},
		"confidence": 100, "order_hint": 1,
	}})
	for want, args := range map[string]M{
		"`attendees` must be a non-empty list":        {"attendees": []any{}, "duration_minutes": 30, "start_window": "a", "end_window": "b"},
		"duration_minutes must be a positive integer": {"attendees": []any{"a@example.com"}, "duration_minutes": 0, "start_window": "a", "end_window": "b"},
	} {
		if msg := callErr(t, cs, "find_meeting_times", args); msg != want {
			t.Fatalf("msg %q", msg)
		}
	}
}
