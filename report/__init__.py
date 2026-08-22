"""Error-report generation: classify findings, plain-language them, render three views.

Promoted out of throwaway session scripts so the recurring-error rules and the
plain-language substitutions live in one place under version control. They had been
re-derived by hand each time, which is how two runs of "the same" report end up
disagreeing about how many distinct kinds of error there are.
"""
from . import classify, plain  # noqa: F401
