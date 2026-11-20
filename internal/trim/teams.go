package trim

// ChatMessage trims a Teams chat or channel message. The snippet comes from
// the body, or from card attachment text when the body has none.
func ChatMessage(raw M, includeBody, includeRaw bool) M {
	body := obj(raw["body"])
	content := str(body["content"])
	contentType := str(body["contentType"])
	user := obj(obj(raw["from"])["user"])

	rawAttachments := list(raw["attachments"])
	attachments := []any{}
	for _, a := range rawAttachments {
		am := obj(a)
		entry := M{
			"id":           am["id"],
			"name":         am["name"],
			"content_type": am["contentType"],
			"content_url":  am["contentUrl"],
		}
		if includeBody {
			entry["content"] = am["content"]
		}
		attachments = append(attachments, entry)
	}
	mentions := []any{}
	for _, m := range list(raw["mentions"]) {
		if t := str(obj(m)["mentionText"]); t != "" {
			mentions = append(mentions, t)
		}
	}
	reactions := M{}
	for _, r := range list(raw["reactions"]) {
		if rt := str(obj(r)["reactionType"]); rt != "" {
			n, _ := reactions[rt].(int)
			reactions[rt] = n + 1
		}
	}
	snippet := Snippet(content, contentType)
	if snippet == nil {
		snippet = CardText(rawAttachments)
	}
	var snippetVal any
	if snippet != nil {
		snippetVal = *snippet
	}
	t := M{
		"id":                  raw["id"],
		"message_type":        or(raw["messageType"], "message"),
		"from":                user["displayName"],
		"from_id":             user["id"],
		"created":             raw["createdDateTime"],
		"last_modified":       raw["lastModifiedDateTime"],
		"deleted":             raw["deletedDateTime"] != nil,
		"importance":          or(raw["importance"], "normal"),
		"subject":             raw["subject"],
		"snippet":             snippetVal,
		"body_type":           or(contentType, "text"),
		"attachments":         attachments,
		"mentions":            mentions,
		"hosted_content_refs": HostedRefs(content),
		"reactions":           reactions,
		"web_url":             raw["webUrl"],
	}
	if includeBody {
		t["body"] = body["content"]
	}
	return withRaw(t, raw, includeRaw)
}

// Chat trims a chat. members are display names (from $expand=members);
// last_message_time is the true last activity, from lastMessagePreview.
func Chat(raw M, includeRaw bool) M {
	members := []any{}
	for _, m := range list(raw["members"]) {
		if name := str(obj(m)["displayName"]); name != "" {
			members = append(members, name)
		}
	}
	return withRaw(M{
		"id":                raw["id"],
		"chat_type":         raw["chatType"],
		"topic":             raw["topic"],
		"members":           members,
		"last_message_time": obj(raw["lastMessagePreview"])["createdDateTime"],
		"metadata_updated":  raw["lastUpdatedDateTime"],
		"web_url":           raw["webUrl"],
	}, raw, includeRaw)
}

// Team trims a joined team.
func Team(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":           raw["id"],
		"display_name": raw["displayName"],
		"description":  raw["description"],
	}, raw, includeRaw)
}

// Channel trims a channel.
func Channel(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":              raw["id"],
		"display_name":    raw["displayName"],
		"description":     raw["description"],
		"membership_type": raw["membershipType"],
		"web_url":         raw["webUrl"],
	}, raw, includeRaw)
}

// HostedDownload trims hosted content returned as base64.
func HostedDownload(raw M, includeRaw bool) M {
	return withRaw(M{
		"content_type":   raw["contentType"],
		"size_bytes":     raw["size"],
		"content_base64": raw["contentBytes"],
	}, raw, includeRaw)
}
