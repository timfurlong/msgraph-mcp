package trim

import "strings"

// Message trims a Graph message.
func Message(raw M, includeBody, includeRaw bool) M {
	body := obj(raw["body"])
	from := raw["from"]
	if !truthy(from) {
		from = raw["sender"]
	}
	t := M{
		"id":                       raw["id"],
		"subject":                  raw["subject"],
		"from":                     email(from),
		"to":                       emails(raw["toRecipients"]),
		"cc":                       emails(raw["ccRecipients"]),
		"received":                 raw["receivedDateTime"],
		"sent":                     raw["sentDateTime"],
		"snippet":                  raw["bodyPreview"],
		"body_type":                or(body["contentType"], "text"),
		"is_read":                  truthy(raw["isRead"]),
		"is_draft":                 truthy(raw["isDraft"]),
		"has_attachments":          truthy(raw["hasAttachments"]),
		"importance":               or(raw["importance"], "normal"),
		"flag":                     or(obj(raw["flag"])["flagStatus"], "notFlagged"),
		"categories":               strList(raw["categories"]),
		"conversation_id":          raw["conversationId"],
		"web_link":                 raw["webLink"],
		"parent_folder_id":         raw["parentFolderId"],
		"inference_classification": raw["inferenceClassification"],
	}
	if includeBody {
		t["body"] = body["content"]
	}
	return withRaw(t, raw, includeRaw)
}

// Folder trims a mail folder.
func Folder(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":                 raw["id"],
		"display_name":       raw["displayName"],
		"parent_folder_id":   raw["parentFolderId"],
		"total_item_count":   raw["totalItemCount"],
		"unread_item_count":  raw["unreadItemCount"],
		"child_folder_count": raw["childFolderCount"],
	}, raw, includeRaw)
}

// AttachmentItem trims an attachment for list_attachments.
func AttachmentItem(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":           raw["id"],
		"name":         raw["name"],
		"content_type": raw["contentType"],
		"size_bytes":   raw["size"],
		"is_inline":    truthy(raw["isInline"]),
		"kind":         AttachmentKind(raw),
	}, raw, includeRaw)
}

// AttachmentKind names an attachment's Graph type: "file", "item" (an
// attached Outlook message, event or contact) or "reference" (a link to a
// cloud file). Any other type comes back as its bare @odata.type name.
func AttachmentKind(raw M) string {
	t, _ := raw["@odata.type"].(string)
	switch t {
	case "#microsoft.graph.fileAttachment":
		return "file"
	case "#microsoft.graph.itemAttachment":
		return "item"
	case "#microsoft.graph.referenceAttachment":
		return "reference"
	case "":
		if _, ok := raw["contentBytes"].(string); ok {
			return "file"
		}
		return "unknown"
	}
	return strings.TrimPrefix(t, "#microsoft.graph.")
}

// ItemType is the bare Graph type of an item attachment's expanded item
// ("message", "event", "eventMessage", "contact"), or nil.
func ItemType(item M) any {
	t, _ := item["@odata.type"].(string)
	if t == "" {
		return nil
	}
	return strings.TrimPrefix(t, "#microsoft.graph.")
}

// AttachedItem trims the expanded item of an item attachment to the fields
// a reader needs: headers and body for a message, the event fields for an
// event, names and addresses for a contact.
func AttachedItem(item M) any {
	if len(item) == 0 {
		return nil
	}
	switch ItemType(item) {
	case "event":
		return Event(item, true, false)
	case "contact":
		addrs := []any{}
		for _, a := range list(item["emailAddresses"]) {
			am := obj(a)
			addrs = append(addrs, M{"name": am["name"], "address": am["address"]})
		}
		return M{
			"display_name":    item["displayName"],
			"email_addresses": addrs,
			"company_name":    item["companyName"],
			"job_title":       item["jobTitle"],
			"mobile_phone":    item["mobilePhone"],
			"business_phones": strList(item["businessPhones"]),
		}
	}
	body := obj(item["body"])
	from := item["from"]
	if !truthy(from) {
		from = item["sender"]
	}
	t := M{
		"subject":         item["subject"],
		"from":            email(from),
		"to":              emails(item["toRecipients"]),
		"cc":              emails(item["ccRecipients"]),
		"sent":            item["sentDateTime"],
		"received":        item["receivedDateTime"],
		"has_attachments": truthy(item["hasAttachments"]),
		"body_type":       or(body["contentType"], "text"),
		"body":            body["content"],
	}
	if mt, ok := item["meetingMessageType"]; ok {
		t["meeting_message_type"] = mt
	}
	return t
}

// AttachmentDownload trims an attachment with its content.
func AttachmentDownload(raw M, includeRaw bool) M {
	return withRaw(M{
		"name":           raw["name"],
		"content_type":   raw["contentType"],
		"size_bytes":     raw["size"],
		"content_base64": raw["contentBytes"],
	}, raw, includeRaw)
}

// User trims /me.
func User(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":                  raw["id"],
		"display_name":        raw["displayName"],
		"user_principal_name": raw["userPrincipalName"],
		"mail":                raw["mail"],
		"job_title":           raw["jobTitle"],
	}, raw, includeRaw)
}

// nonEmpty returns nil for an empty list, as the Python serializer did for
// string-list rule predicates.
func nonEmpty(v any) any {
	if l, ok := v.([]any); ok && len(l) == 0 {
		return nil
	}
	return v
}

func ruleConditions(v any) any {
	c, ok := v.(M)
	if !ok {
		return nil
	}
	return M{
		"sender_contains":          nonEmpty(c["senderContains"]),
		"subject_contains":         nonEmpty(c["subjectContains"]),
		"body_contains":            nonEmpty(c["bodyContains"]),
		"body_or_subject_contains": nonEmpty(c["bodyOrSubjectContains"]),
		"from_addresses":           emails(c["fromAddresses"]),
		"has_attachments":          c["hasAttachments"],
		"categories":               nonEmpty(c["categories"]),
	}
}

func ruleActions(v any) any {
	a, ok := v.(M)
	if !ok {
		return nil
	}
	return M{
		"move_to_folder":        a["moveToFolder"],
		"copy_to_folder":        a["copyToFolder"],
		"delete":                a["delete"],
		"mark_as_read":          a["markAsRead"],
		"mark_importance":       a["markImportance"],
		"assign_categories":     nonEmpty(a["assignCategories"]),
		"forward_to":            emails(a["forwardTo"]),
		"stop_processing_rules": a["stopProcessingRules"],
	}
}

// Rule trims an inbox message rule. Exceptions are left out, as in Python.
func Rule(raw M, includeRaw bool) M {
	return withRaw(M{
		"id":           raw["id"],
		"display_name": raw["displayName"],
		"sequence":     raw["sequence"],
		"is_enabled":   truthy(raw["isEnabled"]),
		"has_error":    truthy(raw["hasError"]),
		"is_read_only": truthy(raw["isReadOnly"]),
		"conditions":   ruleConditions(raw["conditions"]),
		"actions":      ruleActions(raw["actions"]),
	}, raw, includeRaw)
}
