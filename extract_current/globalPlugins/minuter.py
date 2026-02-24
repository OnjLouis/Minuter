# Minuter: second and minute time cues for NVDA
# Copyright (C) 2026 Andre Louis

from __future__ import annotations

import os
import time
import threading
import wave
import audioop
from dataclasses import dataclass
from typing import Optional, List, Tuple

import addonHandler
import config
import globalPluginHandler
import gui
import logHandler
import queueHandler
import speech
import tones

from scriptHandler import script

addonHandler.initTranslation()

try:
	import wx
except Exception:
	wx = None

try:
	import nvwave
except Exception:
	nvwave = None


_CONF_SECTION = "minuter"
_TICK_TRIM_MS = 120  # trim tick.wav to avoid overlap (your tick.wav is 2s)


def _addon_root() -> str:
	return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _sounds_dir() -> str:
	return os.path.join(_addon_root(), "sounds")


def _wav_path(name: str) -> str:
	return os.path.join(_sounds_dir(), name)


def _read_wav_pcm(path: str) -> Optional[Tuple[int, int, int, bytes]]:
	"""Return (channels, sampleRate, bitsPerSample, pcmFrames). Only supports PCM WAV."""
	try:
		with wave.open(path, "rb") as wf:
			ch = int(wf.getnchannels())
			rate = int(wf.getframerate())
			width = int(wf.getsampwidth())  # bytes
			bits = width * 8
			frames = wf.readframes(wf.getnframes())
			return ch, rate, bits, frames
	except Exception:
		return None


def _convert_to_16bit(pcm: bytes, src_bits: int) -> Optional[bytes]:
	"""Convert linear PCM to 16-bit if needed (24-bit -> 16-bit)."""
	try:
		src_width = max(1, src_bits // 8)
		if src_width == 2:
			return pcm
		return audioop.lin2lin(pcm, src_width, 2)
	except Exception:
		return None


def _trim_pcm(pcm: bytes, rate: int, channels: int, bits: int, ms: int) -> bytes:
	width = max(1, bits // 8)
	frame_bytes = channels * width
	target_frames = max(1, int(rate * (ms / 1000.0)))
	target_bytes = target_frames * frame_bytes
	return pcm[: min(len(pcm), target_bytes)]


@dataclass
class _SoundEvent:
	kind: str
	start_time: float
	pcm: bytes
	pos: int = 0


class _AudioMixer(threading.Thread):
	"""
	Software mixer feeding a single nvwave.WavePlayer.
	Converts 24-bit wav -> 16-bit in memory for stable playback.
	Trims tick.wav so it behaves as a tick (no overlap).
	"""

	def __init__(self):
		super().__init__(daemon=True)
		self._stop = threading.Event()
		self._lock = threading.Lock()
		self._events: List[_SoundEvent] = []
		self._player = None

		# Standardize output to 16-bit for NVDA reliability.
		self._ch: int = 2
		self._rate: int = 48000
		self._bits: int = 16
		self._width: int = 2
		self._frame_bytes: int = self._ch * self._width

		self._tick_pcm: Optional[bytes] = None
		self._minute_pcm: Optional[bytes] = None

		self._init_format_and_cache()

	def _init_format_and_cache(self) -> None:
		t = _read_wav_pcm(_wav_path("tick.wav"))
		m = _read_wav_pcm(_wav_path("minute.wav"))

		src = t or m
		if src:
			src_ch, src_rate, _src_bits, _ = src
			self._ch = src_ch
			self._rate = src_rate
			self._bits = 16
			self._width = 2
			self._frame_bytes = self._ch * self._width

		if t and (t[0], t[1]) == (self._ch, self._rate):
			conv = _convert_to_16bit(t[3], t[2])
			if conv is not None:
				self._tick_pcm = _trim_pcm(conv, self._rate, self._ch, self._bits, _TICK_TRIM_MS)

		if m and (m[0], m[1]) == (self._ch, self._rate):
			conv = _convert_to_16bit(m[3], m[2])
			if conv is not None:
				self._minute_pcm = conv

	def start_player(self) -> None:
		if not nvwave:
			return
		try:
			WavePlayer = getattr(nvwave, "WavePlayer", None)
			if WavePlayer is None:
				return
			self._player = WavePlayer(channels=self._ch, samplesPerSec=self._rate, bitsPerSample=self._bits)
		except Exception:
			self._player = None

	def stop(self) -> None:
		self._stop.set()

	def _enqueue(self, kind: str, pcm: bytes) -> None:
		now = time.monotonic()
		with self._lock:
			# Tick should never overlap itself.
			if kind == "tick":
				self._events = [e for e in self._events if e.kind != "tick"]
			self._events.append(_SoundEvent(kind=kind, start_time=now, pcm=pcm, pos=0))

	def queue_tick(self) -> None:
		if self._tick_pcm:
			self._enqueue("tick", self._tick_pcm)
			return
		# Fallback (should not be hit if files are valid)
		if nvwave:
			play_fn = getattr(nvwave, "playWaveFile", None)
			if callable(play_fn):
				try:
					play_fn(_wav_path("tick.wav"))
				except Exception:
					pass

	def queue_minute(self) -> None:
		if self._minute_pcm:
			self._enqueue("minute", self._minute_pcm)
			return
		if nvwave:
			play_fn = getattr(nvwave, "playWaveFile", None)
			if callable(play_fn):
				try:
					play_fn(_wav_path("minute.wav"))
				except Exception:
					pass

	def _prune_and_has_ready(self, now: float) -> bool:
		with self._lock:
			self._events = [e for e in self._events if e.pos < len(e.pcm)]
			return any(e.start_time <= now for e in self._events)

	def _mix_chunk(self, chunk_frames: int, now: float) -> bytes:
		chunk_len = chunk_frames * self._frame_bytes
		out = b"\x00" * chunk_len
		width = self._width

		with self._lock:
			self._events = [e for e in self._events if e.pos < len(e.pcm)]
			active = [e for e in self._events if e.start_time <= now]
			if not active:
				return out

			for e in active:
				remaining = len(e.pcm) - e.pos
				take = chunk_len if remaining >= chunk_len else remaining
				src = e.pcm[e.pos:e.pos + take]
				if take < chunk_len:
					src += b"\x00" * (chunk_len - take)

				try:
					src = audioop.mul(src, width, 0.85)
					out = audioop.add(out, src, width)
				except Exception:
					pass

				e.pos += take

		return out

	def run(self) -> None:
		self.start_player()
		if not self._player:
			return

		chunk_ms = 10
		chunk_frames = max(1, int(self._rate * (chunk_ms / 1000.0)))
		chunk_dur = chunk_frames / float(self._rate)
		next_feed_time = time.monotonic()

		try:
			while not self._stop.is_set():
				now = time.monotonic()

				if not self._prune_and_has_ready(now):
					next_feed_time = time.monotonic()
					self._stop.wait(timeout=0.005)
					continue

				if now < next_feed_time:
					self._stop.wait(timeout=min(0.05, next_feed_time - now))
					continue

				chunk_now = time.monotonic()
				chunk = self._mix_chunk(chunk_frames, chunk_now)
				try:
					self._player.feed(chunk)
				except Exception:
					break

				next_feed_time += chunk_dur

				if (time.monotonic() - next_feed_time) > 0.15:
					next_feed_time = time.monotonic()
		finally:
			try:
				self._player.close()
			except Exception:
				pass


class _SecondWorker(threading.Thread):
	"""Calls onSecond(sec) aligned to wall clock. Debounced by epoch second."""

	def __init__(self, onSecond):
		super().__init__(daemon=True)
		self._onSecond = onSecond
		self._stopEvent = threading.Event()
		self._lastEpochSecond: Optional[int] = None

	def stop(self) -> None:
		self._stopEvent.set()

	def run(self) -> None:
		while not self._stopEvent.is_set():
			now = time.time()
			next_epoch = int(now) + 1
			sleepFor = max(0.0, next_epoch - now)
			if self._stopEvent.wait(timeout=sleepFor):
				break
			try:
				epoch_sec = int(time.time())
				if epoch_sec == self._lastEpochSecond:
					continue
				self._lastEpochSecond = epoch_sec
				sec = time.localtime(epoch_sec).tm_sec
				self._onSecond(sec)
			except Exception:
				pass


class _SettingsDialog(wx.Dialog):
	def __init__(self, parent, plugin: "GlobalPlugin"):
		super().__init__(parent, title=_("Minuter"))
		self._plugin = plugin
		self._didSetInitialFocus = False

		panel = wx.Panel(self)
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		secondsBox = wx.StaticBox(panel, label=_("Second cues"))
		secondsSizer = wx.StaticBoxSizer(secondsBox, wx.VERTICAL)

		minuteBox = wx.StaticBox(panel, label=_("Minute cue"))
		minuteSizer = wx.StaticBoxSizer(minuteBox, wx.VERTICAL)

		self.chkTick = wx.CheckBox(panel, label=_("&Tick every second"))
		self.chkBeepLow = wx.CheckBox(panel, label=_("Beep &low (250 Hz) every second"))
		self.chkWarn = wx.CheckBox(panel, label=_("&Warning beeps at 55–59 (1 kHz)"))
		self.chkSpeak = wx.CheckBox(panel, label=_("&Speak seconds"))
		self.chkMinute = wx.CheckBox(panel, label=_("&Ding at the top of each minute"))

		c = config.conf[_CONF_SECTION]
		self.chkTick.SetValue(bool(c["tickEachSecond"]))
		self.chkBeepLow.SetValue(bool(c["beepEachSecondLow"]))
		self.chkMinute.SetValue(bool(c["dingEachMinute"]))
		self.chkWarn.SetValue(bool(c["beepEndOfMinute"]))
		self.chkSpeak.SetValue(bool(c["speakSeconds"]))

		for chk in (self.chkTick, self.chkBeepLow, self.chkWarn, self.chkSpeak, self.chkMinute):
			chk.Bind(wx.EVT_CHECKBOX, self.onToggle)

		secondsSizer.Add(self.chkTick, flag=wx.ALL, border=6)
		secondsSizer.Add(self.chkBeepLow, flag=wx.ALL, border=6)
		secondsSizer.Add(self.chkWarn, flag=wx.ALL, border=6)
		secondsSizer.Add(self.chkSpeak, flag=wx.ALL, border=6)

		minuteSizer.Add(self.chkMinute, flag=wx.ALL, border=6)

		mainSizer.Add(secondsSizer, flag=wx.EXPAND | wx.ALL, border=8)
		mainSizer.Add(minuteSizer, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=8)

		# IMPORTANT: button must have `panel` as parent (not the dialog),
		# otherwise wx asserts when panel owns the sizer.
		btnRow = wx.BoxSizer(wx.HORIZONTAL)
		self.okButton = wx.Button(panel, wx.ID_OK, label=_("OK"))
		btnRow.AddStretchSpacer(1)
		btnRow.Add(self.okButton, flag=wx.ALL, border=6)
		mainSizer.Add(btnRow, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=8)

		panel.SetSizer(mainSizer)
		mainSizer.Fit(self)

		self.okButton.Bind(wx.EVT_BUTTON, self.onOk)

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.Bind(wx.EVT_SHOW, self.onShow)

		self.CentreOnScreen()

	def onShow(self, evt):
		try:
			if evt.IsShown() and not self._didSetInitialFocus:
				self._didSetInitialFocus = True
				self.chkTick.SetFocus()
		except Exception:
			pass
		evt.Skip()

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			try:
				self.Hide()
			except Exception:
				pass
			try:
				self.EndModal(wx.ID_OK)
			except Exception:
				try:
					self.Destroy()
				except Exception:
					pass
			return
		evt.Skip()

	def onOk(self, evt):
		try:
			self.Hide()
		except Exception:
			pass
		try:
			self.EndModal(wx.ID_OK)
		except Exception:
			try:
				self.Destroy()
			except Exception:
				pass

	def onToggle(self, evt):
		c = config.conf[_CONF_SECTION]
		c["tickEachSecond"] = self.chkTick.GetValue()
		c["beepEachSecondLow"] = self.chkBeepLow.GetValue()
		c["dingEachMinute"] = self.chkMinute.GetValue()
		c["beepEndOfMinute"] = self.chkWarn.GetValue()
		c["speakSeconds"] = self.chkSpeak.GetValue()
		config.conf.save()
		self._plugin._ensureWorkersRunning()
		evt.Skip()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = "Minuter"

	__gestures = {
		"kb:NVDA+windows+backspace": "openMinuterDialog",
	}

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

		if _CONF_SECTION not in config.conf.spec:
			config.conf.spec[_CONF_SECTION] = {
				"tickEachSecond": "boolean(default=False)",
				"beepEachSecondLow": "boolean(default=False)",
				"dingEachMinute": "boolean(default=False)",
				"beepEndOfMinute": "boolean(default=False)",
				"speakSeconds": "boolean(default=False)",
			}
		_ = config.conf[_CONF_SECTION]

		self._worker: Optional[_SecondWorker] = None
		self._mixer: Optional[_AudioMixer] = None
		self._dialog: Optional[_SettingsDialog] = None

		self._ensureWorkersRunning()

	def terminate(self):
		try:
			if self._worker:
				self._worker.stop()
				self._worker = None
		except Exception:
			pass
		try:
			if self._mixer:
				self._mixer.stop()
				self._mixer = None
		except Exception:
			pass
		return super().terminate()

	def _ensureWorkersRunning(self) -> None:
		c = config.conf[_CONF_SECTION]
		enabled = bool(
			c["tickEachSecond"]
			or c["beepEachSecondLow"]
			or c["dingEachMinute"]
			or c["beepEndOfMinute"]
			or c["speakSeconds"]
		)
		needAudio = bool(c["tickEachSecond"] or c["dingEachMinute"])

		if needAudio and not self._mixer:
			self._mixer = _AudioMixer()
			self._mixer.start()
		elif (not needAudio) and self._mixer:
			self._mixer.stop()
			self._mixer = None

		if enabled and not self._worker:
			self._worker = _SecondWorker(self._onSecond)
			self._worker.start()
		elif (not enabled) and self._worker:
			self._worker.stop()
			self._worker = None

	def _onSecond(self, sec: int) -> None:
		c = config.conf[_CONF_SECTION]

		if self._mixer:
			if c["tickEachSecond"]:
				self._mixer.queue_tick()
			if c["dingEachMinute"] and sec == 0:
				self._mixer.queue_minute()

		if c["beepEachSecondLow"]:
			queueHandler.queueFunction(queueHandler.eventQueue, tones.beep, 250, 50)

		if c["beepEndOfMinute"] and sec in (55, 56, 57, 58, 59):
			queueHandler.queueFunction(queueHandler.eventQueue, tones.beep, 1000, 50)

		if c["speakSeconds"]:
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, str(sec))

	@script(description=_("Opens the Minuter settings dialog"))
	def script_openMinuterDialog(self, gesture):
		if wx is None:
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, _("GUI not available"))
			return

		def _show():
			try:
				if self._dialog:
					try:
						if self._dialog.IsShown():
							self._dialog.Raise()
							self._dialog.SetFocus()
							return
					except Exception:
						pass

				self._dialog = _SettingsDialog(gui.mainFrame, self)

				gui.mainFrame.prePopup()
				try:
					self._dialog.CentreOnScreen()
					self._dialog.Raise()
					self._dialog.Show()
					self._dialog.ShowModal()
				finally:
					try:
						gui.mainFrame.postPopup()
					except Exception:
						pass

				try:
					self._dialog.Hide()
				except Exception:
					pass
				try:
					self._dialog.Destroy()
				except Exception:
					pass
				self._dialog = None

			except Exception:
				logHandler.log.exception("Minuter: failed to open settings dialog")
				try:
					queueHandler.queueFunction(
						queueHandler.eventQueue,
						speech.speakMessage,
						_("Could not open Minuter dialog; see NVDA log"),
					)
				except Exception:
					pass

		wx.CallAfter(_show)
