"""
(c) 2015 LinkedIn Corp. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License");?you may not use this file except in compliance with the License. You may obtain a copy of the License at  http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software?distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
"""
from collections import namedtuple

ExchangeEventOrganizer = namedtuple('ExchangeEventOrganizer', ['name', 'email'])
ExchangeEventAttendee = namedtuple('ExchangeEventAttendee', ['name', 'email', 'required'])
ExchangeEventResponse = namedtuple('ExchangeEventResponse', ['name', 'email', 'response', 'last_response', 'required'])


RESPONSE_ACCEPTED = u'Accept'
RESPONSE_DECLINED = u'Decline'
RESPONSE_TENTATIVE = u'Tentative'
RESPONSE_UNKNOWN = u'Unknown'

RESPONSES = [RESPONSE_ACCEPTED, RESPONSE_DECLINED, RESPONSE_TENTATIVE, RESPONSE_UNKNOWN]


class BaseExchangeMessageService(object):

  def __init__(self, service, calendar_id):
    self.service = service
    self.calendar_id = calendar_id

  def event(self, id, *args, **kwargs):
    raise NotImplementedError

  def get_event(self, id):
    raise NotImplementedError

  def new_event(self, **properties):
    raise NotImplementedError


class BaseExchangeMessage(object):
  """
  Implements the EWS element: message
  https://msdn.microsoft.com/en-us/library/aa494306(v=exchg.140).aspx
  """

  _id = None  # Exchange identifier for the message
  _change_key = None  # Exchange requires a second key when updating/deleting the message

  service = None
  calendar_id = None

  mime_content=None
  """Contains the native Multipurpose Internet Mail Extensions (MIME) stream of an object that is represented in base64Binary format."""
  item_id=None
  """Contains the unique identifier and change key of an item in the Exchange store. This property is read-only."""
  parent_folder_id=None
  """Represents the identifier of the parent folder that contains the item or folder. This property is read-only."""
  item_class=None
  """Represents the message class of an item."""
  subject=None
  """Represents the subject for Exchange store items and response objects. The subject is limited to 255 characters."""
  sensitivity=None
  """Indicates the sensitivity level of an item."""
  body=None
  """Represents the actual body content of a message."""
  attachments=None
  """Contains the items or files that are attached to an item in the Exchange store."""
  date_time_received=None
  """Represents the date and time that an item in a mailbox was received."""
  size=None
  """Represents the size in bytes of an item. This property is read-only."""
  categories=None
  """Represents a collection of strings that identify to which categories an item in the mailbox belongs."""
  importance=None
  """Describes the importance of an item."""
  in_reply_to=None
  """Represents the identifier of the item to which this item is a reply."""
  is_submitted=None
  """Indicates whether an item has been submitted to the Outbox default folder."""
  is_draft=None
  """Represents whether an item has not yet been sent."""
  is_from_me=None
  """Indicates whether a user sent an item to him or herself."""
  is_resend=None
  """Indicates whether the item had previously been sent."""
  is_unmodified=None
  """Indicates whether the item has been modified."""
  internet_message_headers=None
  """Represents the collection of all Internet message headers that are contained within an item in a mailbox."""
  date_time_sent=None
  """Represents the date and time that an item in a mailbox was sent."""
  date_time_created=None
  """Represents the date and time that a given item in the mailbox was created."""
  response_objects=None
  """Contains a collection of all the response objects that are associated with an item in the Exchange store."""
  reminder_due_by=None
  """Represents the date and time when the event occurs. This is used by the ReminderMinutesBeforeStart element to determine when the reminder is displayed."""
  reminder_is_set=None
  """Indicates whether a reminder has been set for an item in the Exchange store."""
  reminder_minutes_before_start=None
  """Represents the number of minutes before an event when a reminder is displayed."""
  display_cc=None
  """Represents the display string that is used for the contents of the CC line. This is the concatenated string of all CC recipient display names."""
  display_to=None
  """Represents the display string that is used for the contents of the To box. This is the concatenated string of all To recipient display names."""
  has_attachments=None
  """Represents a property that is set to true if an item has at least one visible attachment. This property is read-only."""
  extended_property=None
  """Identifies extended properties on folders and items."""
  culture=None
  """Represents the culture for a given item in a mailbox."""
  sender=None
  """Identifies the sender of an item."""
  to_recipients=None
  """Contains a set of recipients of a message."""
  cc_recipients=None
  """Represents a collection of recipients that will receive a copy of the message."""
  bcc_recipients=None
  """Represents a collection of recipients to receive a blind carbon copy (Bcc) of an e-mail message."""
  is_read_receipt_requested=None
  """Indicates whether the sender of an item requests a read receipt."""
  is_delivery_receipt_requested=None
  """Indicates whether the sender of an item requests a delivery receipt."""
  conversation_index=None
  """Contains a binary ID that represents the thread to which this message belongs."""
  conversation_topic=None
  """Represents the conversation identifier."""
  from_=None
  """Represents the addressee from whom the message was sent."""
  internet_message_id=None
  """Represents the Internet message identifier of an item."""
  is_read=None
  """Indicates whether a message has been read."""
  is_response_requested=None
  """Indicates whether a response to an e-mail message is requested."""
  references=None
  """Represents the Usenet header that is used to correlate replies with their original messages."""
  reply_to=None
  """Identifies a set of addresses to which replies should be sent."""
  effective_rights=None
  """Contains the client's rights based on the permission settings for the item or folder. This element is read-only."""
  received_by=None
  """Identifies the delegate in a delegate access scenario."""
  received_representing=None
  """Identifies the principal in a delegate access scenario."""
  last_modified_name=None
  """Contains the display name of the last user to modify an item."""
  last_modified_time=None
  """Indicates when an item was last modified."""
  is_associated=None
  """Indicates whether the item is associated with a folder."""
  web_client_read_form_query_string=None
  """Represents a URL to concatenate to the Microsoft Office Outlook Web App endpoint to read an item in Outlook Web App."""
  web_client_edit_form_query_string=None
  """Represents a URL to concatenate to the Outlook Web App endpoint to edit an item in Outlook Web App."""
  conversation_id=None
  """Contains the identifier of an item or conversation."""
  unique_body=None
  """Represents an HTML fragment or plain text which represents the unique body of this conversation."""

  WEEKLY_DAYS = [u'Sunday', u'Monday', u'Tuesday', u'Wednesday', u'Thursday', u'Friday', u'Saturday']

  def __init__(self, service, id=None, calendar_id=u'calendar', xml=None, **kwargs):
    self.service = service
    self.calendar_id = calendar_id

    if xml is not None:
      self._init_from_xml(xml)
    elif id is None:
      self._update_properties(kwargs)
    else:
      self._init_from_service(id)

    self._track_dirty_attributes = True  # magically look for changed attributes

  def _init_from_service(self, id):
    """ Connect to the Exchange service and grab all the properties out of it. """
    raise NotImplementedError

  def _init_from_xml(self, xml):
    """ Using already retrieved XML from Exchange, extract properties out of it. """
    raise NotImplementedError


  def _update_properties(self, properties):
    self._track_dirty_attributes = False
    for key in properties:
      setattr(self, key, properties[key])
    self._track_dirty_attributes = True

  def __setattr__(self, key, value):
    """ Magically track public attributes, so we can track what we need to flush to the Exchange store """
    if self._track_dirty_attributes and not key.startswith(u"_"):
      self._dirty_attributes.add(key)

    object.__setattr__(self, key, value)

  def _reset_dirty_attributes(self):
    self._dirty_attributes = set()

  @property
  def id(self):
    """
    Contains the unique identifier and change key of an item in the Exchange store. This property is read-only.
    see: https://msdn.microsoft.com/en-us/library/aa580234(v=exchg.140).aspx
    """
    return self.id

  @id.setter
  def id(self, id):
    self.id = self._build_resource_dictionary(id)
    self._dirty_attributes.add(u'id')

  @property
  def change_key(self):
    """ **Read-only.** When you change an event, Exchange makes you pass a change key to prevent overwriting a previous version. """
    return self._change_key

  @property
  def mime_content(self):
    """
    Contains the native Multipurpose Internet Mail Extensions (MIME) stream of an object that is represented in base64Binary format.
    see: https://msdn.microsoft.com/en-us/library/aa580801(v=exchg.140).aspx
    """
    return self.mime_content

  @mime_content.setter
  def mime_content(self, mime_content):
    self.mime_content = self._build_resource_dictionary(mime_content)
    self._dirty_attributes.add(u'mime_content')


  @property
  def parent_folder_id(self):
    """
    Represents the identifier of the parent folder that contains the item or folder. This property is read-only.
    see: https://msdn.microsoft.com/en-us/library/aa494327(v=exchg.140).aspx
    """
    return self.parent_folder_id

  @parent_folder_id.setter
  def parent_folder_id(self, parent_folder_id):
    self.parent_folder_id = self._build_resource_dictionary(parent_folder_id)
    self._dirty_attributes.add(u'parent_folder_id')

  @property
  def item_class(self):
    """
    Represents the message class of an item.
    see: https://msdn.microsoft.com/en-us/library/aa580993(v=exchg.140).aspx
    """
    return self.item_class

  @item_class.setter
  def item_class(self, item_class):
    self.item_class = self._build_resource_dictionary(item_class)
    self._dirty_attributes.add(u'item_class')

  @property
  def subject(self):
    """
    Represents the subject for Exchange store items and response objects. The subject is limited to 255 characters.
    see: https://msdn.microsoft.com/en-us/library/aa565100(v=exchg.140).aspx
    """
    return self.subject

  @subject.setter
  def subject(self, subject):
    self.subject = self._build_resource_dictionary(subject)
    self._dirty_attributes.add(u'subject')

  @property
  def sensitivity(self):
    """
    Indicates the sensitivity level of an item.
    see: https://msdn.microsoft.com/en-us/library/aa565687(v=exchg.140).aspx
    """
    return self.sensitivity

  @sensitivity.setter
  def sensitivity(self, sensitivity):
    self.sensitivity = self._build_resource_dictionary(sensitivity)
    self._dirty_attributes.add(u'sensitivity')

  @property
  def body(self):
    """
    Represents the actual body content of a message.
    see: https://msdn.microsoft.com/en-us/library/aa581015(v=exchg.140).aspx
    """
    return self.body

  @body.setter
  def body(self, body):
    self.body = self._build_resource_dictionary(body)
    self._dirty_attributes.add(u'body')

  @property
  def attachments(self):
    """
    Contains the items or files that are attached to an item in the Exchange store.
    see: https://msdn.microsoft.com/en-us/library/aa564869(v=exchg.140).aspx
    """
    return self.attachments

  @attachments.setter
  def attachments(self, attachments):
    self.attachments = self._build_resource_dictionary(attachments)
    self._dirty_attributes.add(u'attachments')

  @property
  def date_time_received(self):
    """
    Represents the date and time that an item in a mailbox was received.
    see: https://msdn.microsoft.com/en-us/library/aa564021(v=exchg.140).aspx
    """
    return self.date_time_received

  @date_time_received.setter
  def date_time_received(self, date_time_received):
    self.date_time_received = self._build_resource_dictionary(date_time_received)
    self._dirty_attributes.add(u'date_time_received')

  @property
  def size(self):
    """
    Represents the size in bytes of an item. This property is read-only.
    see: https://msdn.microsoft.com/en-us/library/aa564235(v=exchg.140).aspx
    """
    return self.size


  @property
  def categories(self):
    """
    Represents a collection of strings that identify to which categories an item in the mailbox belongs.
    see: https://msdn.microsoft.com/en-us/library/aa565683(v=exchg.140).aspx
    """
    return self.categories

  @categories.setter
  def categories(self, categories):
    self.categories = self._build_resource_dictionary(categories)
    self._dirty_attributes.add(u'categories')

  @property
  def importance(self):
    """
    Describes the importance of an item.
    see: https://msdn.microsoft.com/en-us/library/aa563467(v=exchg.140).aspx
    """
    return self.importance

  @importance.setter
  def importance(self, importance):
    self.importance = self._build_resource_dictionary(importance)
    self._dirty_attributes.add(u'importance')

  @property
  def in_reply_to(self):
    """
    Represents the identifier of the item to which this item is a reply.
    see: https://msdn.microsoft.com/en-us/library/aa580994(v=exchg.140).aspx
    """
    return self.in_reply_to

  @in_reply_to.setter
  def in_reply_to(self, in_reply_to):
    self.in_reply_to = self._build_resource_dictionary(in_reply_to)
    self._dirty_attributes.add(u'in_reply_to')

  @property
  def is_submitted(self):
    """
    Indicates whether an item has been submitted to the Outbox default folder.
    see: https://msdn.microsoft.com/en-us/library/aa494303(v=exchg.140).aspx
    """
    return self.is_submitted

  @is_submitted.setter
  def is_submitted(self, is_submitted):
    self.is_submitted = self._build_resource_dictionary(is_submitted)
    self._dirty_attributes.add(u'is_submitted')

  @property
  def is_draft(self):
    """
    Represents whether an item has not yet been sent.
    see: https://msdn.microsoft.com/en-us/library/aa581576(v=exchg.140).aspx
    """
    return self.is_draft

  @is_draft.setter
  def is_draft(self, is_draft):
    self.is_draft = self._build_resource_dictionary(is_draft)
    self._dirty_attributes.add(u'is_draft')

  @property
  def is_from_me(self):
    """
    Indicates whether a user sent an item to him or herself.
    see: https://msdn.microsoft.com/en-us/library/aa565618(v=exchg.140).aspx
    """
    return self.is_from_me

  @is_from_me.setter
  def is_from_me(self, is_from_me):
    self.is_from_me = self._build_resource_dictionary(is_from_me)
    self._dirty_attributes.add(u'is_from_me')

  @property
  def is_resend(self):
    """
    Indicates whether the item had previously been sent.
    see: https://msdn.microsoft.com/en-us/library/aa564024(v=exchg.140).aspx
    """
    return self.is_resend

  @is_resend.setter
  def is_resend(self, is_resend):
    self.is_resend = self._build_resource_dictionary(is_resend)
    self._dirty_attributes.add(u'is_resend')

  @property
  def is_unmodified(self):
    """
    Indicates whether the item has been modified.
    see: https://msdn.microsoft.com/en-us/library/aa565038(v=exchg.140).aspx
    """
    return self.is_unmodified

  @is_unmodified.setter
  def is_unmodified(self, is_unmodified):
    self.is_unmodified = self._build_resource_dictionary(is_unmodified)
    self._dirty_attributes.add(u'is_unmodified')

  @property
  def internet_message_headers(self):
    """
    Represents the collection of all Internet message headers that are contained within an item in a mailbox.
    see: https://msdn.microsoft.com/en-us/library/aa580788(v=exchg.140).aspx
    """
    return self.internet_message_headers

  @internet_message_headers.setter
  def internet_message_headers(self, internet_message_headers):
    self.internet_message_headers = self._build_resource_dictionary(internet_message_headers)
    self._dirty_attributes.add(u'internet_message_headers')

  @property
  def date_time_sent(self):
    """
    Represents the date and time that an item in a mailbox was sent.
    see:
    """
    return self.date_time_sent

  @date_time_sent.setter
  def date_time_sent(self, date_time_sent):
    self.date_time_sent = self._build_resource_dictionary(date_time_sent)
    self._dirty_attributes.add(u'date_time_sent')

  @property
  def date_time_created(self):
    """
    Represents the date and time that a given item in the mailbox was created.
    see: https://msdn.microsoft.com/en-us/library/aa580543(v=exchg.140).aspx
    """
    return self.date_time_created

  @date_time_created.setter
  def date_time_created(self, date_time_created):
    self.date_time_created = self._build_resource_dictionary(date_time_created)
    self._dirty_attributes.add(u'date_time_created')

  @property
  def response_objects(self):
    """
    Contains a collection of all the response objects that are associated with an item in the Exchange store.
    see: https://msdn.microsoft.com/en-us/library/aa564717(v=exchg.140).aspx
    """
    return self.response_objects

  @response_objects.setter
  def response_objects(self, response_objects):
    self.response_objects = self._build_resource_dictionary(response_objects)
    self._dirty_attributes.add(u'response_objects')

  @property
  def reminder_due_by(self):
    """
    Represents the date and time when the event occurs. This is used by the ReminderMinutesBeforeStart element to determine when the reminder is displayed.
    see: https://msdn.microsoft.com/en-us/library/aa565894(v=exchg.140).aspx
    """
    return self.reminder_due_by

  @reminder_due_by.setter
  def reminder_due_by(self, reminder_due_by):
    self.reminder_due_by = self._build_resource_dictionary(reminder_due_by)
    self._dirty_attributes.add(u'reminder_due_by')

  @property
  def reminder_is_set(self):
    """
    Indicates whether a reminder has been set for an item in the Exchange store.
    see: https://msdn.microsoft.com/en-us/library/aa566410(v=exchg.140).aspx
    """
    return self.reminder_is_set

  @reminder_is_set.setter
  def reminder_is_set(self, reminder_is_set):
    self.reminder_is_set = self._build_resource_dictionary(reminder_is_set)
    self._dirty_attributes.add(u'reminder_is_set')

  @property
  def reminder_minutes_before_start(self):
    """
    Represents the number of minutes before an event when a reminder is displayed.
    see: https://msdn.microsoft.com/en-us/library/aa581305(v=exchg.140).aspx
    """
    return self.reminder_minutes_before_start

  @reminder_minutes_before_start.setter
  def reminder_minutes_before_start(self, reminder_minutes_before_start):
    self.reminder_minutes_before_start = self._build_resource_dictionary(reminder_minutes_before_start)
    self._dirty_attributes.add(u'reminder_minutes_before_start')

  @property
  def display_cc(self):
    """
    Represents the display string that is used for the contents of the CC line. This is the concatenated string of all CC recipient display names.
    see: https://msdn.microsoft.com/en-us/library/aa564744(v=exchg.140).aspx
    """
    return self.display_cc

  @display_cc.setter
  def display_cc(self, display_cc):
    self.display_cc = self._build_resource_dictionary(display_cc)
    self._dirty_attributes.add(u'display_cc')

  @property
  def display_to(self):
    """
    Represents the display string that is used for the contents of the To box. This is the concatenated string of all To recipient display names.
    see: https://msdn.microsoft.com/en-us/library/aa493834(v=exchg.140).aspx
    """
    return self.display_to

  @display_to.setter
  def display_to(self, display_to):
    self.display_to = self._build_resource_dictionary(display_to)
    self._dirty_attributes.add(u'display_to')

  @property
  def has_attachments(self):
    """
    Represents a property that is set to true if an item has at least one visible attachment. This property is read-only.
    see: https://msdn.microsoft.com/en-us/library/aa580961(v=exchg.140).aspx
    """
    return self.has_attachments

  @has_attachments.setter
  def has_attachments(self, has_attachments):
    self.has_attachments = self._build_resource_dictionary(has_attachments)
    self._dirty_attributes.add(u'has_attachments')

  @property
  def extended_property(self):
    """
    Identifies extended properties on folders and items.
    see: https://msdn.microsoft.com/en-us/library/aa566405(v=exchg.140).aspx
    """
    return self.extended_property

  @extended_property.setter
  def extended_property(self, extended_property):
    self.extended_property = self._build_resource_dictionary(extended_property)
    self._dirty_attributes.add(u'extended_property')

  @property
  def culture(self):
    """
    Represents the culture for a given item in a mailbox.
    see: https://msdn.microsoft.com/en-us/library/aa563712(v=exchg.140).aspx
    """
    return self.culture

  @culture.setter
  def culture(self, culture):
    self.culture = self._build_resource_dictionary(culture)
    self._dirty_attributes.add(u'culture')

  @property
  def sender(self):
    """
    Identifies the sender of an item.
    see: https://msdn.microsoft.com/en-us/library/aa579529(v=exchg.140).aspx
    """
    return self.sender

  @sender.setter
  def sender(self, sender):
    self.sender = self._build_resource_dictionary(sender)
    self._dirty_attributes.add(u'sender')

  @property
  def to_recipients(self):
    """
    Contains a set of recipients of a message.
    see: https://msdn.microsoft.com/en-us/library/aa563719(v=exchg.140).aspx
    """
    return self.to_recipients

  @to_recipients.setter
  def to_recipients(self, to_recipients):
    self.to_recipients = self._build_resource_dictionary(to_recipients)
    self._dirty_attributes.add(u'to_recipients')

  @property
  def cc_recipients(self):
    """
    Represents a collection of recipients that will receive a copy of the message.
    see: https://msdn.microsoft.com/en-us/library/aa581076(v=exchg.140).aspx
    """
    return self.cc_recipients

  @cc_recipients.setter
  def cc_recipients(self, cc_recipients):
    self.cc_recipients = self._build_resource_dictionary(cc_recipients)
    self._dirty_attributes.add(u'cc_recipients')

  @property
  def bcc_recipients(self):
    """
    Represents a collection of recipients to receive a blind carbon copy (Bcc) of an e-mail message.
    see: https://msdn.microsoft.com/en-us/library/aa565250(v=exchg.140).aspx
    """
    return self.bcc_recipients

  @bcc_recipients.setter
  def bcc_recipients(self, bcc_recipients):
    self.bcc_recipients = self._build_resource_dictionary(bcc_recipients)
    self._dirty_attributes.add(u'bcc_recipients')

  @property
  def is_read_receipt_requested(self):
    """
    Indicates whether the sender of an item requests a read receipt.
    see: https://msdn.microsoft.com/en-us/library/aa563919(v=exchg.140).aspx
    """
    return self.is_read_receipt_requested

  @is_read_receipt_requested.setter
  def is_read_receipt_requested(self, is_read_receipt_requested):
    self.is_read_receipt_requested = self._build_resource_dictionary(is_read_receipt_requested)
    self._dirty_attributes.add(u'is_read_receipt_requested')

  @property
  def is_delivery_receipt_requested(self):
    """
    Indicates whether the sender of an item requests a delivery receipt.
    see: https://msdn.microsoft.com/en-us/library/aa564249(v=exchg.140).aspx
    """
    return self.is_delivery_receipt_requested

  @is_delivery_receipt_requested.setter
  def is_delivery_receipt_requested(self, is_delivery_receipt_requested):
    self.is_delivery_receipt_requested = self._build_resource_dictionary(is_delivery_receipt_requested)
    self._dirty_attributes.add(u'is_delivery_receipt_requested')

  @property
  def conversation_index(self):
    """
    Contains a binary ID that represents the thread to which this message belongs.
    see: https://msdn.microsoft.com/en-us/library/aa566462(v=exchg.140).aspx
    """
    return self.conversation_index

  @conversation_index.setter
  def conversation_index(self, conversation_index):
    self.conversation_index = self._build_resource_dictionary(conversation_index)
    self._dirty_attributes.add(u'conversation_index')

  @property
  def conversation_topic(self):
    """
    Represents the conversation identifier.
    see: https://msdn.microsoft.com/en-us/library/aa580685(v=exchg.140).aspx
    """
    return self.conversation_topic

  @conversation_topic.setter
  def conversation_topic(self, conversation_topic):
    self.conversation_topic = self._build_resource_dictionary(conversation_topic)
    self._dirty_attributes.add(u'conversation_topic')

  @property
  def from_(self):
    """
    Represents the addressee from_ whom the message was sent.
    see: https://msdn.microsoft.com/en-us/library/aa581049(v=exchg.140).aspx
    """
    return self.from_

  @from_.setter
  def from_(self, from_):
    self.from_ = self._build_resource_dictionary(from_)
    self._dirty_attributes.add(u'from_')

  @property
  def internet_message_id(self):
    """
    Represents the Internet message identifier of an item.
    see: https://msdn.microsoft.com/en-us/library/aa564528(v=exchg.140).aspx
    """
    return self.internet_message_id

  @internet_message_id.setter
  def internet_message_id(self, internet_message_id):
    self.internet_message_id = self._build_resource_dictionary(internet_message_id)
    self._dirty_attributes.add(u'internet_message_id')

  @property
  def is_read(self):
    """
    Indicates whether a message has been read.
    see: https://msdn.microsoft.com/en-us/library/aa493829(v=exchg.140).aspx
    """
    return self.is_read

  @is_read.setter
  def is_read(self, is_read):
    self.is_read = self._build_resource_dictionary(is_read)
    self._dirty_attributes.add(u'is_read')

  @property
  def is_response_requested(self):
    """
    Indicates whether a response to an e-mail message is requested.
    see: https://msdn.microsoft.com/en-us/library/aa563990(v=exchg.140).aspx
    """
    return self.is_response_requested

  @is_response_requested.setter
  def is_response_requested(self, is_response_requested):
    self.is_response_requested = self._build_resource_dictionary(is_response_requested)
    self._dirty_attributes.add(u'is_response_requested')

  @property
  def references(self):
    """
    Represents the Usenet header that is used to correlate replies with their original messages.
    see: https://msdn.microsoft.com/en-us/library/aa565671(v=exchg.140).aspx
    """
    return self.references

  @references.setter
  def references(self, references):
    self.references = self._build_resource_dictionary(references)
    self._dirty_attributes.add(u'references')

  @property
  def reply_to(self):
    """
    Identifies a set of addresses to which replies should be sent.
    see: https://msdn.microsoft.com/en-us/library/aa563522(v=exchg.140).aspx
    """
    return self.reply_to

  @reply_to.setter
  def reply_to(self, reply_to):
    self.reply_to = self._build_resource_dictionary(reply_to)
    self._dirty_attributes.add(u'reply_to')

  @property
  def effective_rights(self):
    """
    Contains the client's rights based on the permission settings for the item or folder. This element is read-only.
    see: https://msdn.microsoft.com/en-us/library/bb891883(v=exchg.140).aspx
    """
    return self.effective_rights

  @effective_rights.setter
  def effective_rights(self, effective_rights):
    self.effective_rights = self._build_resource_dictionary(effective_rights)
    self._dirty_attributes.add(u'effective_rights')

  @property
  def received_by(self):
    """
    Identifies the delegate in a delegate access scenario.
    see: https://msdn.microsoft.com/en-us/library/bb891874(v=exchg.140).aspx
    """
    return self.received_by

  @received_by.setter
  def received_by(self, received_by):
    self.received_by = self._build_resource_dictionary(received_by)
    self._dirty_attributes.add(u'received_by')

  @property
  def received_representing(self):
    """
    Identifies the principal in a delegate access scenario.
    see: https://msdn.microsoft.com/en-us/library/bb891812(v=exchg.140).aspx
    """
    return self.received_representing

  @received_representing.setter
  def received_representing(self, received_representing):
    self.received_representing = self._build_resource_dictionary(received_representing)
    self._dirty_attributes.add(u'received_representing')

  @property
  def last_modified_name(self):
    """
    Contains the display name of the last user to modify an item.
    see: https://msdn.microsoft.com/en-us/library/bb891829(v=exchg.140).aspx
    """
    return self.last_modified_name

  @last_modified_name.setter
  def last_modified_name(self, last_modified_name):
    self.last_modified_name = self._build_resource_dictionary(last_modified_name)
    self._dirty_attributes.add(u'last_modified_name')

  @property
  def last_modified_time(self):
    """
    Indicates when an item was last modified.
    see: https://msdn.microsoft.com/en-us/library/bb891845(v=exchg.140).aspx
    """
    return self.last_modified_time

  @last_modified_time.setter
  def last_modified_time(self, last_modified_time):
    self.last_modified_time = self._build_resource_dictionary(last_modified_time)
    self._dirty_attributes.add(u'last_modified_time')

  @property
  def is_associated(self):
    """
    Indicates whether the item is associated with a folder.
    see: https://msdn.microsoft.com/en-us/library/dd899429(v=exchg.140).aspx
    """
    return self.is_associated

  @is_associated.setter
  def is_associated(self, is_associated):
    self.is_associated = self._build_resource_dictionary(is_associated)
    self._dirty_attributes.add(u'is_associated')

  @property
  def web_client_read_form_query_string(self):
    """
    Represents a URL to concatenate to the Microsoft Office Outlook Web App endpoint to read an item in Outlook Web App.
    see: https://msdn.microsoft.com/en-us/library/dd877102(v=exchg.140).aspx
    """
    return self.web_client_read_form_query_string

  @web_client_read_form_query_string.setter
  def web_client_read_form_query_string(self, web_client_read_form_query_string):
    self.web_client_read_form_query_string = self._build_resource_dictionary(web_client_read_form_query_string)
    self._dirty_attributes.add(u'web_client_read_form_query_string')

  @property
  def web_client_edit_form_query_string(self):
    """
    Represents a URL to concatenate to the Outlook Web App endpoint to edit an item in Outlook Web App.
    see: https://msdn.microsoft.com/en-us/library/dd899477(v=exchg.140).aspx
    """
    return self.web_client_edit_form_query_string

  @web_client_edit_form_query_string.setter
  def web_client_edit_form_query_string(self, web_client_edit_form_query_string):
    self.web_client_edit_form_query_string = self._build_resource_dictionary(web_client_edit_form_query_string)
    self._dirty_attributes.add(u'web_client_edit_form_query_string')

  @property
  def conversation_id(self):
    """
    Contains the identifier of an item or conversation.
    see: https://msdn.microsoft.com/en-us/library/dd899527(v=exchg.140).aspx
    """
    return self.conversation_id

  @conversation_id.setter
  def conversation_id(self, conversation_id):
    self.conversation_id = self._build_resource_dictionary(conversation_id)
    self._dirty_attributes.add(u'conversation_id')

  @property
  def unique_body(self):
    """
    Represents an HTML fragment or plain text which represents the unique body of this conversation.
    see: https://msdn.microsoft.com/en-us/library/dd877075(v=exchg.140).aspx
    """
    return self.unique_body

  @unique_body.setter
  def unique_body(self, unique_body):
    self.unique_body = self._build_resource_dictionary(unique_body)
    self._dirty_attributes.add(u'unique_body')

