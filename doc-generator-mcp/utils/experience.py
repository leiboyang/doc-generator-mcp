"""
经验记录模块 —— 大模型驱动的经验知识库（含准入机制）

工作流 (方向B):
1. 任务前: 大模型调用 get_experience 查询相关历史教训
2. 任务中: 大模型参考历史教训，避免重复犯错
3. 任务后: 大模型调用 log_experience 回写新发现的教训

准入机制:
- 每条经验必须有来源(source)和证据(evidence)
- source=verified: 经过实际执行验证的事实，直接入库
- source=user_feedback: 用户明确指出的问题，直接入库
- source=inferred: 大模型推理的结论，标记 needs_review，查询时默认不返回
- 写入前自动去重：检查是否已有高度相似的条目
- 通过 review_experience 可审批/驳回待审条目

存储格式: JSONL (每日一个文件)
"""

import json
import os
import threading
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional
from difflib import SequenceMatcher


# 来源可信度
SOURCE_TRUST = {
    "verified": 3,       # 实际执行验证
    "user_feedback": 3,  # 用户明确反馈
    "inferred": 1,       # 大模型推理（待验证）
}

# 相似度阈值（超过此值视为重复）
SIMILARITY_THRESHOLD = 0.75


class ExperienceLogger:
    """经验记录器"""

    def __init__(self, experience_dir: str = None):
        if experience_dir is None:
            experience_dir = str(Path(__file__).parent.parent / "experiences")
        self.experience_dir = Path(experience_dir)
        self.experience_dir.mkdir(parents=True, exist_ok=True)

    def log(
        self,
        category: str,
        lessons: list[str],
        source: str = "inferred",
        evidence: str = "",
        summary: str = "",
        context: str = "",
        importance: str = "normal",
    ) -> dict:
        """记录经验教训（含准入检查）

        Args:
            category: 经验分类
            lessons: 教训列表
            source: 来源 — "verified"(执行验证) / "user_feedback"(用户反馈) / "inferred"(推理)
            evidence: 证据（错误信息、执行结果、用户原话等）
            summary: 操作摘要
            context: 背景描述
            importance: "high" / "normal" / "low"

        Returns:
            记录结果（含去重/待审信息）
        """
        # 参数校验
        if source not in SOURCE_TRUST:
            return {
                "success": False,
                "error": f"无效来源: {source}，必须是 verified/user_feedback/inferred",
            }

        if not lessons:
            return {"success": False, "error": "lessons 不能为空"}

        if not evidence and source != "inferred":
            return {
                "success": False,
                "error": f"来源为 {source} 时必须提供 evidence（证据）",
            }

        # 去重检查
        duplicates = self._find_similar(lessons, category, days=90)
        if duplicates:
            new_lessons = []
            for lesson in lessons:
                is_dup = False
                for dup in duplicates:
                    for existing_lesson in dup.get("lessons", []):
                        sim = SequenceMatcher(None, lesson.lower(), existing_lesson.lower()).ratio()
                        if sim > SIMILARITY_THRESHOLD:
                            is_dup = True
                            break
                    if is_dup:
                        break
                if not is_dup:
                    new_lessons.append(lesson)

            if not new_lessons:
                return {
                    "success": False,
                    "error": "所有教训与已有经验重复",
                    "duplicates": [d["summary"] for d in duplicates],
                }

            lessons = new_lessons

        # 确定是否需要审核
        needs_review = (source == "inferred")

        entry = {
            "timestamp": datetime.now().isoformat(),
            "category": category,
            "summary": summary,
            "context": context,
            "lessons": lessons,
            "source": source,
            "evidence": evidence,
            "importance": importance,
            "needs_review": needs_review,
        }

        today = datetime.now().strftime("%Y%m%d")
        file_path = self.experience_dir / f"experience_{today}.jsonl"
        with open(file_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

        status = "pending_review" if needs_review else "accepted"
        return {
            "success": True,
            "status": status,
            "category": category,
            "lessons_count": len(lessons),
            "duplicates_removed": len(duplicates),
            "file": str(file_path),
        }

    def review(
        self,
        action: str,
        entry_id: str = None,
        category: str = None,
    ) -> dict:
        """审核待审经验

        Args:
            action: "approve"(批准) / "reject"(驳回) / "list"(列出待审)
            entry_id: 指定条目的时间戳ID（可选）
            category: 按分类过滤（可选）

        Returns:
            审核结果
        """
        if action == "list":
            pending = self.query(needs_review=True, days=365)
            if category:
                pending = [e for e in pending if e.get("category") == category]
            return {
                "success": True,
                "pending_count": len(pending),
                "entries": pending,
            }

        if action in ("approve", "reject"):
            # 找到对应条目并更新
            updated = 0
            today = datetime.now()

            for i in range(365):
                date = (today - timedelta(days=i)).strftime("%Y%m%d")
                file_path = self.experience_dir / f"experience_{date}.jsonl"
                if not file_path.exists():
                    continue

                lines = []
                with open(file_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            entry = json.loads(line)
                        except json.JSONDecodeError:
                            lines.append(line)
                            continue

                        # 匹配条目
                        ts = entry.get("timestamp", "")
                        match = False
                        if entry_id and ts == entry_id:
                            match = True
                        elif not entry_id and entry.get("needs_review"):
                            if not category or entry.get("category") == category:
                                match = True

                        if match and entry.get("needs_review"):
                            if action == "approve":
                                entry["needs_review"] = False
                                entry["reviewed_at"] = datetime.now().isoformat()
                                entry["review_result"] = "approved"
                            else:
                                entry["reviewed_at"] = datetime.now().isoformat()
                                entry["review_result"] = "rejected"
                            updated += 1

                        lines.append(json.dumps(entry, ensure_ascii=False))

                # 原子写入：先写临时文件，再 rename 覆盖原文件
                if updated > 0:
                    tmp_path = file_path.with_suffix(".jsonl.tmp")
                    with open(tmp_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(lines) + "\n")
                    tmp_path.replace(file_path)

                # 如果是指定单条目审核，找到后提前退出
                if entry_id and updated > 0:
                    break

            return {"success": True, "updated": updated, "action": action}

        return {"success": False, "error": f"未知操作: {action}"}

    def query(
        self,
        category: str = None,
        keyword: str = None,
        days: int = 30,
        importance: str = None,
        needs_review: bool = None,
        include_pending: bool = False,
    ) -> list[dict]:
        """查询经验（默认排除待审条目）

        Args:
            category: 按分类过滤
            keyword: 按关键词搜索
            days: 回溯天数
            importance: 按重要度过滤
            needs_review: 只查待审条目 (True) / 只查已审条目 (False) / 不过滤 (None)
            include_pending: 是否包含待审条目（默认 False，防止污染）

        Returns:
            匹配的经验条目列表
        """
        entries = []
        today = datetime.now()

        for i in range(min(days, 365)):
            date = (today - timedelta(days=i)).strftime("%Y%m%d")
            file_path = self.experience_dir / f"experience_{date}.jsonl"
            if file_path.exists():
                with open(file_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            entry = json.loads(line)
                        except json.JSONDecodeError:
                            continue

                        # 跳过已驳回的条目
                        if entry.get("review_result") == "rejected":
                            continue

                        # 待审过滤
                        is_pending = entry.get("needs_review", False)
                        if needs_review is not None:
                            if is_pending != needs_review:
                                continue
                        elif is_pending and not include_pending:
                            # 默认排除待审条目
                            continue

                        # 分类过滤
                        if category and entry.get("category") != category:
                            continue

                        # 重要度过滤
                        if importance and entry.get("importance") != importance:
                            continue

                        # 关键词搜索
                        if keyword:
                            kw_lower = keyword.lower()
                            searchable = " ".join([
                                " ".join(entry.get("lessons", [])),
                                entry.get("summary", ""),
                                entry.get("context", ""),
                                entry.get("category", ""),
                                entry.get("evidence", ""),
                            ]).lower()
                            if kw_lower not in searchable:
                                continue

                        entries.append(entry)

        return entries

    def get_all_categories(self, days: int = 90) -> list[dict]:
        """列出所有经验分类（仅统计已审核条目）"""
        entries = self.query(days=days, include_pending=False)
        pending = self.query(days=days, needs_review=True)
        cat_counts = {}
        for e in entries:
            cat = e.get("category", "uncategorized")
            if cat not in cat_counts:
                cat_counts[cat] = {"count": 0, "high_importance": 0, "latest": "", "pending": 0}
            cat_counts[cat]["count"] += 1
            if e.get("importance") == "high":
                cat_counts[cat]["high_importance"] += 1
            ts = e.get("timestamp", "")
            if ts > cat_counts[cat]["latest"]:
                cat_counts[cat]["latest"] = ts

        for e in pending:
            cat = e.get("category", "uncategorized")
            if cat in cat_counts:
                cat_counts[cat]["pending"] += 1
            else:
                cat_counts[cat] = {"count": 0, "high_importance": 0, "latest": "", "pending": 1}

        return [
            {"category": cat, **info}
            for cat, info in sorted(cat_counts.items(), key=lambda x: x[1]["count"], reverse=True)
        ]

    def _find_similar(self, lessons: list[str], category: str, days: int = 90) -> list[dict]:
        """查找相似条目（去重用）"""
        existing = self.query(category=category, days=days, include_pending=True)
        similar = []
        for entry in existing:
            for existing_lesson in entry.get("lessons", []):
                for new_lesson in lessons:
                    sim = SequenceMatcher(None, new_lesson.lower(), existing_lesson.lower()).ratio()
                    if sim > SIMILARITY_THRESHOLD:
                        similar.append(entry)
                        break
                else:
                    continue
                break
        return similar


# 全局实例（线程安全）
_logger: Optional[ExperienceLogger] = None
_logger_lock = threading.Lock()


def get_experience_logger(experience_dir: str = None) -> ExperienceLogger:
    """获取全局经验记录器实例（线程安全）"""
    global _logger
    if _logger is not None:
        return _logger
    with _logger_lock:
        if _logger is None:
            _logger = ExperienceLogger(experience_dir)
    return _logger
