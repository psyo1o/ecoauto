# -*- coding: utf-8 -*-
"""GUI 취소 버튼용 threading.Event 헬퍼."""


def is_cancelled(cancel_event) -> bool:
    return cancel_event is not None and cancel_event.is_set()
