package trim

// Calendar trims a calendar.
func Calendar(raw M, includeRaw bool) M {
	var owner any
	if o := obj(raw["owner"]); len(o) > 0 {
		owner = M{"name": o["name"], "address": o["address"]}
	}
	return withRaw(M{
		"id":                  raw["id"],
		"name":                raw["name"],
		"owner":               owner,
		"can_edit":            truthy(raw["canEdit"]),
		"is_default_calendar": truthy(raw["isDefaultCalendar"]),
	}, raw, includeRaw)
}

// Event trims a calendar event. recurrence is Graph's recurrence object.
func Event(raw M, includeBody, includeRaw bool) M {
	start, end := obj(raw["start"]), obj(raw["end"])
	attendees := []any{}
	for _, a := range list(raw["attendees"]) {
		am := obj(a)
		var name, address any
		if e, ok := email(am).(M); ok {
			name, address = e["name"], e["address"]
		}
		attendees = append(attendees, M{
			"name":     name,
			"address":  address,
			"type":     am["type"],
			"response": obj(am["status"])["response"],
		})
	}
	t := M{
		"id":                            raw["id"],
		"subject":                       raw["subject"],
		"organizer":                     email(raw["organizer"]),
		"start":                         M{"date_time": start["dateTime"], "time_zone": start["timeZone"]},
		"end":                           M{"date_time": end["dateTime"], "time_zone": end["timeZone"]},
		"location":                      obj(raw["location"])["displayName"],
		"is_all_day":                    truthy(raw["isAllDay"]),
		"is_cancelled":                  truthy(raw["isCancelled"]),
		"is_online_meeting":             truthy(raw["isOnlineMeeting"]),
		"attendees":                     attendees,
		"body_preview":                  raw["bodyPreview"],
		"show_as":                       raw["showAs"],
		"sensitivity":                   raw["sensitivity"],
		"is_reminder_on":                raw["isReminderOn"],
		"reminder_minutes_before_start": raw["reminderMinutesBeforeStart"],
		"recurrence":                    raw["recurrence"],
		"web_link":                      raw["webLink"],
	}
	if truthy(raw["isOnlineMeeting"]) {
		t["online_meeting_url"] = obj(raw["onlineMeeting"])["joinUrl"]
	}
	if includeBody {
		t["body"] = obj(raw["body"])["content"]
	}
	return withRaw(t, raw, includeRaw)
}
