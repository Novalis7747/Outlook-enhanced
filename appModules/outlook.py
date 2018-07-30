# -*- coding: UTF-8 -*-
# Outlook: a NVDA appModule for Outlook Messagefields
#Copyright (C) 2018 Ralf Kefferpuetz, other contributors
# Released under GPL 2

import api
import appModuleHandler
import controlTypes
import NVDAObjects
from nvdaBuiltin.appModules import outlook
from NVDAObjects.IAccessible import IAccessible
import ui
import winUser
import windowUtils
import tones
import scriptHandler

"""
appModule for Microsoft Outlook, Version 0.9
Written by Ralf Kefferpuetz, 2018
Tested under Outlook 2016 (32bit Office suits) -this addOn my work in older versions of Outlook, like Outlook 2013 but it is only supported in Outlook 2016 32bit
Features:
Hotkey Alt-1 to Alt-8 to read the following fields in a mail or RSS-message:
- From, Sent, To, CC, BCC, Subject, Infobar and number and name of Attachments
-Single press: speaks the field content
- double press: moves focus to the field
- trible press: copies content of that field to the clipboard (except of Alt-8 for Attachments)
Upcoming:
-your suggestion please
"""

class AppModule(outlook.AppModule):

	def script_from(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4097)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			# for RSS mails
			try:
				handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4107)
				if handle:
					# found handle
					w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
					ui.message(" %s %s" % (w.name, w.value))
			except LookupError:
				#checking if in composing mode, then put focus to the From button
				try:
					handle = windowUtils.findDescendantWindow(fg.windowHandle, className="Button", controlID=4257)
					if handle:
						w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
						if scriptHandler.getLastScriptRepeatCount() == 1:
							winUser.setForegroundWindow(handle)
						elif scriptHandler.getLastScriptRepeatCount() == 0:
							ui.message(" %s %s" % (w.name, w.value))
				except LookupError:
					tones.beep(440, 20)
	# Translators: Documentation for from script.
	script_from.__doc__=_("speaks the From field, double press moves the focus to the From field, trible press copies the From field value to the clipboard.")
  
	def script_sent(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4098)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			# for RSS mails
			try:
				handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4105)
				if handle:
					# found handle
					w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
					ui.message(" %s %s" % (w.name, w.value))
			except LookupError:
				tones.beep(440, 20)
	# Translators: Documentation for Sent script.
	script_sent.__doc__=_("speaks the Sent field, double press moves the focus to the Sent field, trible press copies the Sent field value to the clipboard.")

	def script_to(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4099)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			tones.beep(440, 20)
	# Translators: Documentation for to script.
	script_to.__doc__=_("speaks the To field, double press moves the focus to the To field, trible press copies the To field value to the clipboard.")

	def script_cc(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4100)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			if api.getForegroundObject().appModule.productName	== "Microsoft Outlook":
				tones.beep(440, 20)
			else:
				tones.beep(440, 20)
	# Translators: Documentation for CC script.
	script_cc.__doc__=_("speaks the CC field, double press moves the focus to the CC field, trible press copies the CC field value to the clipboard.")

	def script_bcc(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4103)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			tones.beep(440, 20)
	# Translators: Documentation for BCC script.
	script_bcc.__doc__=_("speaks the BCC field, double press moves the focus to the BCC field, trible press copies the BCC field value to the clipboard.")

	def script_subject(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4101)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			# for RSS mails
			try:
				handle = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4108)
				if handle:
					# found handle
					w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
					ui.message(" %s %s" % (w.name, w.value))
			except LookupError:
				tones.beep(440, 20)
	# Translators: Documentation for Subject script.
	script_subject.__doc__=_("speaks the Subject field, double press moves the focus to the Subject field, trible press copies the Subject field value to the clipboard.")

	def script_infobar(self, gesture):
		global orig
		fg = api.getForegroundObject()
		try:
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="rctrl_renwnd32", controlID=4262)
			handleSubject = windowUtils.findDescendantWindow(fg.windowHandle, className="RichEdit20WPT", controlID=4101)
			if handle and handleSubject:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				if scriptHandler.getLastScriptRepeatCount() == 2:
					# trible press, copy to clipboard and set focus to original field
					api.copyToClip(w.value)
					ui.message(_("Copied to clipboard"))
					api.setNavigatorObject(w,isFocus=True)
					orig.setFocus()
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, w.value))
					orig = api.getFocusObject()
		except LookupError:
			tones.beep(440, 20)
	# Translators: Documentation for infoBar script.
	script_infobar.__doc__=_("speaks the InfoBar field, double press moves the focus to the InfoBar field, trible press copies the InfoBar field value to the clipboard.")

	def script_attachments(self, gesture):
		fg = api.getForegroundObject()
		try:
			obj = api.getFocusObject()
			appName = obj.appModule.productName
			appVersion = obj.appModule.productVersion
			if appVersion.startswith('15.'):
				ui.message("use tab or shift-tab to go to the attachments in Outlook 2013")
				tones.beep(440, 20)
				return
			handle = windowUtils.findDescendantWindow(fg.windowHandle, className="rctrl_renwnd32", controlID=4306)
			if handle:
				# found handle
				w = NVDAObjects.IAccessible.getNVDAObjectFromEvent(handle, winUser.OBJID_CLIENT, 0)
				try:
					wc = w.firstChild.firstChild.firstChild.firstChild.childCount
					indexString = (" %s " % (wc))
					children = w.firstChild.firstChild.firstChild.firstChild.children
					for child in children:
						ui.message(child.name)
				except:
					indexString = 0
					pass
				if scriptHandler.getLastScriptRepeatCount() == 1:
				# double press, set focus in field
					winUser.setForegroundWindow(handle)
				else:
					# single press
					ui.message(" %s %s" % (w.name, indexString))
		except LookupError:
			tones.beep(440, 20)
	# Translators: Documentation for Attachments script.
	script_attachments.__doc__=_("speaks the number and name of the Attachments, double press moves the focus to the Attachments.")

	__gestures={
		"kb:alt+1":"from",
		"kb:alt+2":"sent",
		"kb:alt+3":"to",
		"kb:alt+4":"cc",
	"kb:alt+5":"bcc",
		"kb:alt+6":"subject",
		"kb:alt+8":"attachments",
		"kb:alt+7":"infobar"
	}
