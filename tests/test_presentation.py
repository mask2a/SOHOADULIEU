import queue
import unittest
from unittest.mock import patch

from sohoa.domain import Progress
from sohoa.presentation import App


class Widget:
    def __init__(self, mode='indeterminate'):
        self.options = {'mode': mode}

    def configure(self, **options):
        self.options.update(options)

    def __getitem__(self, key):
        return self.options[key]

    def start(self):
        pass

    def stop(self):
        pass


class Root:
    def after(self, *_args):
        pass


class PresentationTests(unittest.TestCase):
    def test_completion_is_handled_even_with_backlogged_progress_events(self):
        app = App.__new__(App)
        app.events = queue.Queue()
        app.progress_events = queue.Queue(maxsize=1)
        app.results = queue.Queue()
        for _ in range(500):
            app.events.put(Progress('Đang đọc dữ liệu'))
        app.results.put(('done', 'output'))
        app.busy = True
        app.root = Root()
        app.progress = Widget()
        app.merge_btn = Widget()
        app.compare_btn = Widget()
        app.compare_batch_btn = Widget()
        app.cancel_btn = Widget()
        app.status = Widget()
        app.progress_percent = Widget()
        app.progress_percent.set = lambda _value: None
        logged = []
        app.log = logged.append

        with patch('sohoa.presentation.messagebox.showinfo'):
            app.poll()

        self.assertFalse(app.busy)
        self.assertEqual(app.cancel_btn.options['state'], 'disabled')
        self.assertEqual(app.status.options['text'], 'Đã xuất kết quả')
        self.assertTrue(app.events.empty())
        self.assertEqual(logged, ['Đã xuất kết quả: output'])

    def test_progress_percentage_updates_status_without_filling_log(self):
        app = App.__new__(App)
        app.events = queue.Queue()
        app.progress_events = queue.Queue(maxsize=1)
        app.progress_events.put(Progress('Đang đọc tệp: school.xlsx', 25, 100, log_message=False))
        app.results = queue.Queue()
        app.busy = True
        app.root = Root()
        app.progress = Widget()
        app.merge_btn = Widget()
        app.compare_btn = Widget()
        app.compare_batch_btn = Widget()
        app.cancel_btn = Widget()
        app.status = Widget()
        app.progress_percent = Widget()
        percentages = []
        app.progress_percent.set = percentages.append
        logged = []
        app.log = logged.append

        app.poll()

        self.assertEqual(app.status.options['text'], 'Đang đọc tệp: school.xlsx')
        self.assertEqual(app.progress.options['value'], 25)
        self.assertEqual(app.progress.options['maximum'], 100)
        self.assertEqual(percentages, ['25.0%'])
        self.assertEqual(logged, [])

    def test_latest_progress_replaces_stale_progress(self):
        app = App.__new__(App)
        app.events = queue.Queue()
        app.progress_events = queue.Queue(maxsize=1)
        app.results = queue.Queue()
        app.emit_progress(Progress('Đang đọc cũ', 18, 100, log_message=False))
        app.emit_progress(Progress('Đang đọc mới', 19, 100, log_message=False))

        latest = app.progress_events.get_nowait()

        self.assertEqual(latest.message, 'Đang đọc mới')
        self.assertEqual(latest.current, 19)
        self.assertTrue(app.events.empty())
