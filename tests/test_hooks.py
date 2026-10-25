#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : test_hooks.py
"""Hook 事件调度器（hooks.py）单元测试。"""

from wordformat.hooks import HookRegistry


class TestRegisterEmit:
    def test_emit_without_callbacks_returns_data(self):
        reg = HookRegistry()
        data = reg.emit("before_node_format", node="n")
        assert data == {"node": "n"}

    def test_emit_merges_callback_return(self):
        reg = HookRegistry()
        reg.register("on_config_loaded", lambda d: {"config": d["config"] + 1}, name="cb")
        result = reg.emit("on_config_loaded", config=1)
        assert result["config"] == 2

    def test_callback_can_mutate_event_data(self):
        reg = HookRegistry()

        def cb(data):
            data["issues"].append("x")

        reg.register("on_caption_check", cb)
        result = reg.emit("on_caption_check", issues=["a"])
        assert result["issues"] == ["a", "x"]

    def test_callback_exception_caught_others_continue(self):
        reg = HookRegistry()

        def bad(_d):
            raise RuntimeError("boom")

        reg.register("on_summary_build", bad, name="bad")
        reg.register("on_summary_build", lambda d: {"summary": "ok"}, name="good")
        result = reg.emit("on_summary_build", check=True)
        assert result["summary"] == "ok"

    def test_callback_return_none_keeps_data(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="noop")
        result = reg.emit("on_comment", text="t")
        assert result["text"] == "t"

    def test_register_non_callable_ignored(self):
        reg = HookRegistry()
        reg.register("on_comment", 123)  # 非可调用对象：警告后忽略
        assert not reg.is_registered("on_comment")

    def test_callbacks_run_in_registration_order(self):
        reg = HookRegistry()
        order = []
        reg.register("on_comment", lambda d: order.append(1), name="a")
        reg.register("on_comment", lambda d: order.append(2), name="b")
        reg.emit("on_comment")
        assert order == [1, 2]

    def test_callback_without_name_uses_function_name(self):
        reg = HookRegistry()

        def my_callback(_d):
            return None

        reg.register("on_comment", my_callback)
        names = [n for n, _cb in reg.registered_callbacks("on_comment")]
        assert names == ["my_callback"]


class TestUnregister:
    def test_unregister_by_name(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="a")
        reg.register("on_comment", lambda d: None, name="b")
        assert reg.unregister("on_comment", "a") is True
        names = [n for n, _cb in reg.registered_callbacks("on_comment")]
        assert names == ["b"]
        # 再次移除同名的返回 False
        assert reg.unregister("on_comment", "a") is False

    def test_unregister_name_none_clears_all(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="a")
        assert reg.unregister("on_comment", None) is True
        assert not reg.is_registered("on_comment")
        # 空事件再清空返回 False
        assert reg.unregister("on_comment", None) is False

    def test_unregister_missing_event_returns_false(self):
        reg = HookRegistry()
        assert reg.unregister("on_comment", "nope") is False


class TestInspect:
    def test_registered_events_and_callbacks(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="x")
        reg.register("on_comment", lambda d: None, name="y")
        reg.register("on_caption_check", lambda d: None, name="z")
        assert set(reg.registered_events()) == {"on_comment", "on_caption_check"}
        assert len(reg.registered_callbacks()) == 3
        assert len(reg.registered_callbacks("on_comment")) == 2

    def test_has_handlers_tracks_registration(self):
        reg = HookRegistry()
        assert reg.has_handlers is False
        reg.register("on_comment", lambda d: None, name="x")
        assert reg.has_handlers is True
        reg.unregister("on_comment", "x")
        assert reg.has_handlers is False

    def test_clear(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="x")
        reg.clear()
        assert reg.registered_events() == []
        assert reg.has_handlers is False

    def test_preset_loaded_state_reset_by_clear(self):
        reg = HookRegistry()
        reg.mark_preset_loaded("p")
        assert reg.is_preset_loaded("p") is True
        reg.clear()
        assert reg.is_preset_loaded("p") is False
