# -*- coding: utf-8 -*-
"""验证「所选打印机生效」: 打印期间默认打印机切换 + 作业真实进入目标打印机队列.

手法: 暂停目标打印机 (不出纸) -> 发送作业 -> 检查队列 -> 清除作业 -> 恢复打印机。
"""
import os
import tempfile
import time

import win32print

import printer_core as pc


def jobs(printer):
    try:
        h = win32print.OpenPrinter(printer)
        try:
            return [j["pDocument"] for j in win32print.EnumJobs(h, 0, -1, 1)]
        finally:
            win32print.ClosePrinter(h)
    except Exception as e:
        return [f"<无法查询: {e}>"]


def job_ids(printer):
    """(作业ID, 文档名) 列表"""
    try:
        h = win32print.OpenPrinter(printer)
        try:
            return [(j["JobId"], j["pDocument"]) for j in win32print.EnumJobs(h, 0, -1, 1)]
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return []


def delete_job(printer, job_id):
    """只删除指定作业 (绝不清空整个队列)"""
    h = win32print.OpenPrinter(printer, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
    try:
        win32print.SetJob(h, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
    finally:
        win32print.ClosePrinter(h)


def control(printer, cmd):
    h = win32print.OpenPrinter(printer, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
    try:
        win32print.SetPrinter(h, 0, None, cmd)
    finally:
        win32print.ClosePrinter(h)


def main():
    default = pc.get_default_printer_name()
    allp = [p[2] for p in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    others = [p for p in allp if p != default]
    print("默认打印机:", default)
    print("候选目标:", others)

    target = next((p for p in others if "DocuCentre" in p), others[0])
    print("测试目标打印机:", repr(target))

    td = tempfile.mkdtemp()
    txt = os.path.join(td, "target_test.txt")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("BatchPrinter 打印机选择测试\n" + "x\n" * 5)

    paused = False
    try:
        control(target, win32print.PRINTER_CONTROL_PAUSE)
        paused = True
        print("已暂停目标打印机 (作业会留在队列, 不出纸)")
    except Exception as e:
        print("无法暂停目标打印机:", e)

    logs = []

    try:
        print("\n=== 1) default_printer_guard 切换/恢复 ===")
        with pc.default_printer_guard(target, logs.append) as g:
            during = pc.get_default_printer_name()
            print("  guard 内默认打印机:", during, "| changed =", g.changed)
            assert during == target, f"切换失败: {during} != {target}"
        after = pc.get_default_printer_name()
        print("  guard 后默认打印机:", after)
        assert after == default, f"未恢复: {after} != {default}"
        print("  ✅ 切换与恢复都正确")

        print("\n=== 2) GDI 通道 (文本) 打到非默认打印机 ===")
        before_t, before_d = jobs(target), jobs(default)
        pc.print_file(txt, target, {"copies": 1, "color": "mono", "duplex": None, "pages": None},
                      {}, "", logs.append)
        time.sleep(2)
        after_t, after_d = jobs(target), jobs(default)
        print("  目标队列 前/后:", before_t, "->", after_t)
        print("  默认队列 前/后:", before_d, "->", after_d)
        assert len(after_t) > len(before_t), "作业没有进入目标打印机队列"
        assert len(after_d) == len(before_d), "默认打印机队列不应增加"
        print("  ✅ 文本作业进入了所选打印机")

        print("\n=== 3) Shell 通道 (print 动词, 典型 Office 走法) ===")
        before_t, before_d = jobs(target), jobs(default)
        logs2 = []
        try:
            pc._shell_print_with_defaults(txt, target, None, pc.DMCOLOR_MONOCHROME, None, logs2.append)
            time.sleep(3)
            after_t, after_d = jobs(target), jobs(default)
            print("  目标队列:", before_t, "->", after_t)
            print("  默认队列:", before_d, "->", after_d)
            print("  日志:", logs2)
            if len(after_t) > len(before_t):
                print("  ✅ Shell 打印也打到了所选打印机")
            else:
                print("  ⚠ Shell 通道未产生队列作业 (记事本可能未注册 print 动词), 未见打到默认打印机")
                assert len(after_d) == len(before_d), "作业错误地打到了默认打印机!"
        except Exception as e:
            print("  Shell 通道异常:", e)
        print("\n  Shell 执行后默认打印机:", pc.get_default_printer_name())
        assert pc.get_default_printer_name() == default, "默认打印机未恢复"

    finally:
        # 只删除本次测试产生的作业, 绝不动队列里已有的其他作业
        try:
            for pid, pname in job_ids(target):
                if "target_test" in str(pname):
                    delete_job(target, pid)
                    print("\n已删除测试作业:", pname)
        except Exception as e:
            print("删除测试作业失败:", e)
        if paused:
            try:
                control(target, win32print.PRINTER_CONTROL_RESUME)
                print("已恢复目标打印机")
            except Exception as e:
                print("恢复打印机失败:", e)

    print("\n最终默认打印机:", pc.get_default_printer_name(), "(应为", default, ")")
    print("打印日志:\n  " + "\n  ".join(logs))
    print("\n✅ 打印机选择链路验证完成")


if __name__ == "__main__":
    main()
