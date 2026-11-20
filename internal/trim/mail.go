package trim

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
	}, raw, includeRaw)
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
