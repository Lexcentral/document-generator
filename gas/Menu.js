function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Worship')
    .addItem('Check Benevolence Figures',        'runBenevolenceCheck')
    .addSeparator()
    .addItem('Generate Staff Agenda',           'showStaffAgendaDialog')
    .addItem('Generate Deacon Agenda',           'showDeaconAgendaDialog')
    .addItem('Generate Council Packet',          'showCouncilPacketDialog')
    .addItem('Generate Business Meeting Packet', 'showBusinessPacketDialog')
    .addSeparator()
    .addItem('Generate Leadership Schedule…',  'showLeadershipScheduleDialog')
    .addItem('Send Attendee Reminder…',        'showAttendeeReminderDialog')
    .addItem('Send Test Reminder Email…',      'showTestReminderDialog')
    .addItem('Install Reminder Trigger',       'installReminderTrigger')
    .addToUi();
}
