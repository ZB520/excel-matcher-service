import re
from pathlib import Path
from typing import Set, Tuple

import pandas as pd
from rapidfuzz import fuzz


# ========== 用户需手动修改的路径 ==========

# 旧表路径（包含列：学校, 序号, 班级人数, 书号, 书名, 出版社, 单价, 学生数量）
OLD_FILE_PATH = r"C:\Users\ZB\Desktop\cursor project\汽车高级技工-机械系.xlsx"

# 新表路径（包含列：书号, 书名, 出版社, 单价, 数量）
NEW_FILE_PATH = r"C:\Users\ZB\Desktop\cursor project\汽车技工.xlsx"

# 输出文件名（生成在当前工作目录下）
OUTPUT_MATCHED_PATH = r"C:\Users\ZB\Desktop\cursor project\matched_details.xlsx"
OUTPUT_UNMATCHED_PATH = r"C:\Users\ZB\Desktop\cursor project\unmatched_items.xlsx"
OUTPUT_MATCHED_ORIGINAL_PATH = r"C:\Users\ZB\Desktop\cursor project\matched_original_raw.xlsx"
LOG_PATH = r"match_log.txt"

# 书名相似度阈值（0-100，越高越严格）
TITLE_SIM_THRESHOLD = 75

# 学科关键词（用于互斥，避免不同学科之间误匹配）
SUBJECT_KEYWORDS = [
    "语文",
    "英语",
    "数学",
    "信息技术",
    "化学",
    "物理",
    "道德与法治",
    "心理与健康",
    "历史",
    "地理",
]

# 册别关键词（用于区分上册/下册等版本）
VOLUME_KEYWORDS = ["上册", "下册"]


# ========== 工具函数 ==========

def normalize_title(value) -> str:
    """
    书名强标准化：
    - 转为字符串、小写
    - 去除前后空格
    - 统一空白字符
    - 删除所有空格和常见标点/符号
    """
    if pd.isna(value):
        return ""

    s = str(value)
    s = s.strip().lower()
    # 全角空格替换为半角空格
    s = s.replace("\u3000", " ")
    # 把各种空白（包括制表符等）合并为一个空格
    s = re.sub(r"\s+", " ", s)

    # 去掉标题前面的年份前缀（如 2025、2024），避免不同年份版本难以匹配
    s = re.sub(r"^(19|20)\d{2}", "", s)

    # 需要删除的字符（标点 + 空格）
    remove_chars = (
        " ，,。.:：；;!！?？-—_（）()【】[]『』《》<>“”\"'、 "
    )
    trans = str.maketrans("", "", remove_chars)
    s = s.translate(trans)

    # 去除常见“噪音”修饰词（不影响学科和册别），只保留课程核心名称
    noise_phrases = [
        "省编一体化",
        "省编",
        "一体化",
        "规划教材",
        "精品教材",
        "中等职业教育",
        "国家级",
        "十三五",
        "含微课",
        "new",
        "全彩",
        "双色",
        "中高职一体教材",
        "宁波版",
        "彩绘",
        "100%折扣",
        "*",
        "**",
        "***",
        "****",
        "*****",
        "微课版",
        "新课标",
        "修订版",
        "折扣100%",
    ]
    for phrase in noise_phrases:
        s = s.replace(phrase, "")

    # 去掉“第X版”这类版本信息（数字或常见汉字数字）
    s = re.sub(r"第[0-9一二三四五六七八九十]+版", "", s)

    # 再次规整空白并去除首尾空格
    s = re.sub(r"\s+", " ", s).strip()

    return s


def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """去除列名前后空格，并转为字符串。"""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_and_prepare_old(path: str) -> Tuple[pd.DataFrame, Tuple[str, ...]]:
    """读取并预处理旧表，返回 DataFrame 和原始列顺序。"""
    df = pd.read_excel(path)
    df = _standardize_columns(df)

    required_cols = ["学校", "序号", "班级人数", "书号", "书名", "出版社", "单价", "学生数量"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"旧表缺少必要列: {missing}，当前列为: {list(df.columns)}")

    original_columns = tuple(df.columns)

    # 数值列处理
    df["学生数量"] = pd.to_numeric(df["学生数量"], errors="coerce").fillna(0)
    df["单价"] = pd.to_numeric(df["单价"], errors="coerce")

    # 书名标准化
    df["norm_title"] = df["书名"].map(normalize_title)

    return df, original_columns


def load_and_prepare_new(path: str) -> pd.DataFrame:
    """读取并预处理新表。"""
    df = pd.read_excel(path)
    df = _standardize_columns(df)

    required_cols = ["书号", "书名", "出版社", "单价", "数量"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"新表缺少必要列: {missing}，当前列为: {list(df.columns)}")

    # 数值列处理
    df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0)
    df["单价"] = pd.to_numeric(df["单价"], errors="coerce")

    # 书名标准化
    df["norm_title"] = df["书名"].map(normalize_title)

    return df


def aggregate_old(df_old: pd.DataFrame) -> pd.DataFrame:
    """按标准化书名聚合旧表学生数量。"""
    grouped = (
        df_old.groupby("norm_title", dropna=False)["学生数量"]
        .sum()
        .reset_index(name="old_total_students")
    )
    return grouped


def aggregate_new(df_new: pd.DataFrame) -> pd.DataFrame:
    """按标准化书名聚合新表数量。"""
    grouped = (
        df_new.groupby("norm_title", dropna=False)["数量"]
        .sum()
        .reset_index(name="new_total_quantity")
    )
    return grouped


def build_match_sets(
    old_group: pd.DataFrame, new_group: pd.DataFrame
) -> Tuple[pd.DataFrame, Set[str], Set[str], Set[str]]:
    """
    基于书名相似度构建匹配成功/数量冲突/仅新表书名集合，并返回匹配明细表。

    - 使用标准化后的 norm_title 做相似度比较
    - 对每个旧表 norm_title，在新表中寻找相似度最高的书名
    - 当相似度 >= TITLE_SIM_THRESHOLD 时认为是同一本书
    """
    matches = []

    def _extract_keyword(title: str, candidates) -> str | None:
        """从规范化书名中提取第一个出现的关键词。"""
        for kw in candidates:
            if kw in title:
                return kw
        return None

    for _, old_row in old_group.iterrows():
        old_title = old_row["norm_title"]
        old_total = old_row["old_total_students"]

        best_score = -1.0
        best_new_title = None
        best_new_total = None

        for _, new_row in new_group.iterrows():
            new_title = new_row["norm_title"]

            # 学科关键词互斥：如果两边都包含学科且学科不同，则不考虑匹配
            old_subject = _extract_keyword(old_title, SUBJECT_KEYWORDS)
            new_subject = _extract_keyword(new_title, SUBJECT_KEYWORDS)
            if old_subject and new_subject and old_subject != new_subject:
                continue

            # 册别关键词互斥：如果两边都标明上册/下册且不同，则不考虑匹配
            old_volume = _extract_keyword(old_title, VOLUME_KEYWORDS)
            new_volume = _extract_keyword(new_title, VOLUME_KEYWORDS)
            if old_volume and new_volume and old_volume != new_volume:
                continue

            score = float(fuzz.ratio(old_title, new_title))
            if score > best_score:
                best_score = score
                best_new_title = new_title
                best_new_total = new_row["new_total_quantity"]

        if best_new_title is None or best_score < TITLE_SIM_THRESHOLD:
            continue

        status = "matched" if old_total == best_new_total else "conflict"
        matches.append(
            {
                "old_norm_title": old_title,
                "new_norm_title": best_new_title,
                "similarity": best_score,
                "old_total_students": old_total,
                "new_total_quantity": best_new_total,
                "status": status,
            }
        )

    if not matches:
        empty_df = pd.DataFrame(
            columns=[
                "old_norm_title",
                "new_norm_title",
                "similarity",
                "old_total_students",
                "new_total_quantity",
                "status",
            ]
        )
        return empty_df, set(), set(), set(new_group["norm_title"])

    matches_df = pd.DataFrame(matches)

    # 一个新表标准化书名只允许匹配到一个旧表标准化书名：
    # 先按相似度从高到低排序，再对 new_norm_title 去重，保留每个新表书名的最佳匹配
    matches_df = matches_df.sort_values(by="similarity", ascending=False)
    matches_df = matches_df.drop_duplicates(subset=["new_norm_title"], keep="first")

    matched_titles: Set[str] = set(
        matches_df.loc[matches_df["status"] == "matched", "old_norm_title"]
    )
    conflict_titles: Set[str] = set(
        matches_df.loc[matches_df["status"] == "conflict", "old_norm_title"]
    )

    matched_new_titles: Set[str] = set(matches_df["new_norm_title"])
    new_only_titles: Set[str] = set(new_group["norm_title"]) - matched_new_titles

    return matches_df, matched_titles, conflict_titles, new_only_titles


def build_newinfo_mapping(df_new: pd.DataFrame, matches_df: pd.DataFrame) -> pd.DataFrame:
    """
    为匹配成功的书名构建新表信息映射：
    - 先按数量降序，再按原顺序，去重保留每个 norm_title 一行。
    """
    if matches_df.empty:
        return pd.DataFrame(
            columns=[
                "norm_title",
                "书号",
                "书名",
                "出版社",
                "单价",
                "匹配相似度",
                "新表标准化书名",
                "_new_order",
            ]
        )

    # 给新表记录一个行顺序索引，以便后续按新表顺序排序
    df_new_with_order = df_new.copy()
    df_new_with_order["_new_order"] = range(len(df_new_with_order))

    # 将匹配结果映射回新表，取每个旧表书名对应的新表信息
    subset = df_new_with_order.merge(
        matches_df[["old_norm_title", "new_norm_title", "similarity"]],
        left_on="norm_title",
        right_on="new_norm_title",
        how="inner",
    )

    # 以旧表标准化书名作为主键，方便与旧表合并
    subset["norm_title"] = subset["old_norm_title"]
    subset.rename(
        columns={
            "similarity": "匹配相似度",
            "new_norm_title": "新表标准化书名",
        },
        inplace=True,
    )

    # 先按新表原始顺序、再按数量降序，优先保留新表中靠前且数量较大的记录
    subset = subset.sort_values(by=["_new_order", "数量"], ascending=[True, False])

    mapping = subset.drop_duplicates(subset=["norm_title"])[
        [
            "norm_title",
            "书号",
            "书名",
            "出版社",
            "单价",
            "匹配相似度",
            "新表标准化书名",
            "_new_order",
        ]
    ].copy()

    return mapping


def apply_replacements(
    df_old: pd.DataFrame,
    mapping_df: pd.DataFrame,
    matched_titles: Set[str],
    original_columns: Tuple[str, ...],
) -> pd.DataFrame:
    """在旧表上替换指定字段，生成匹配明细表。"""
    matched_old = df_old[df_old["norm_title"].isin(matched_titles)].copy()
    # 记录旧表中的行顺序，便于在每个书名分组内保持原始顺序
    matched_old["_old_order"] = matched_old.index

    merged = matched_old.merge(mapping_df, on="norm_title", how="left", suffixes=("", "_new"))

    for col in ["书号", "书名", "出版社", "单价"]:
        new_col = f"{col}_new"
        if new_col in merged.columns:
            merged[col] = merged[new_col].where(merged[new_col].notna(), merged[col])
            merged.drop(columns=[new_col], inplace=True)

    # 按新表书名顺序排序；同一书名内按旧表原始行顺序排序
    if "_new_order" in merged.columns and "_old_order" in merged.columns:
        merged = merged.sort_values(by=["_new_order", "_old_order"])

    # 删除辅助列，按旧表原始列顺序输出
    if "norm_title" in merged.columns:
        merged.drop(columns=["norm_title"], inplace=True)
    for helper_col in ["_new_order", "_old_order"]:
        if helper_col in merged.columns:
            merged.drop(columns=[helper_col], inplace=True)

    # original_columns 中可能已经不包含 norm_title，这里安全过滤一次
    ordered_cols = [c for c in original_columns if c in merged.columns]
    # 保证不遗漏后续新增列
    remaining_cols = [c for c in merged.columns if c not in ordered_cols]
    result = merged[ordered_cols + remaining_cols]
    return result


def build_unmatched_table(
    df_new: pd.DataFrame, matches_df: pd.DataFrame, new_only_titles: Set[str]
) -> pd.DataFrame:
    """基于冲突书名和仅新表书名生成未匹配表，并增加状态列。"""
    unmatched_new_titles: Set[str] = set(new_only_titles)

    status_map = {}
    if not matches_df.empty:
        conflict_new_titles = set(
            matches_df.loc[matches_df["status"] == "conflict", "new_norm_title"]
        )
        unmatched_new_titles |= conflict_new_titles
        status_map = {
            row["new_norm_title"]: row["status"]
            for _, row in matches_df.iterrows()
            if row["status"] == "conflict"
        }

    unmatched = df_new[df_new["norm_title"].isin(unmatched_new_titles)].copy()

    if not unmatched.empty:
        def _status(norm_title: str) -> str:
            if norm_title in status_map:
                return "数量冲突"
            return "仅新表"

        unmatched["状态"] = unmatched["norm_title"].map(_status)

    # 保留 norm_title 方便后续人工检查
    return unmatched


def write_log(
    old_path: str,
    new_path: str,
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    matched_titles: Set[str],
    conflict_titles: Set[str],
    new_only_titles: Set[str],
    matched_path: str,
    unmatched_path: str,
    matched_original_path: str,
) -> None:
    """把关键信息写入日志文件。"""
    lines = []
    lines.append(f"旧表路径: {old_path}")
    lines.append(f"新表路径: {new_path}")
    lines.append(f"旧表总行数: {len(df_old)}")
    lines.append(f"新表总行数: {len(df_new)}")
    lines.append(f"书名相似度阈值: {TITLE_SIM_THRESHOLD}")
    lines.append(f"匹配成功的书名数量: {len(matched_titles)}")
    lines.append(f"数量不一致（冲突）的书名数量: {len(conflict_titles)}")
    lines.append(f"仅出现在新表中的书名数量: {len(new_only_titles)}")
    lines.append(f"已匹配数据表输出文件: {matched_path}")
    lines.append(f"未匹配数据表输出文件: {unmatched_path}")
    lines.append(f"已匹配数据原始表输出文件: {matched_original_path}")

    log_text = "\n".join(lines)

    try:
        Path(LOG_PATH).write_text(log_text, encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        print(f"写入日志文件失败: {exc}")

    print(log_text)


def run_matching(
    old_path: str | Path,
    new_path: str | Path,
    matched_path: str | Path,
    unmatched_path: str | Path,
    matched_original_path: str | Path,
) -> dict[str, Path]:
    """
    运行一次完整的匹配流程。

    参数均为路径字符串或 Path 对象，返回值为输出结果文件的路径字典，键包括：
    - matched_details
    - unmatched_items
    - matched_original
    """
    old_path = Path(old_path)
    new_path = Path(new_path)
    matched_path = Path(matched_path)
    unmatched_path = Path(unmatched_path)
    matched_original_path = Path(matched_original_path)

    df_old, original_columns = load_and_prepare_old(str(old_path))
    df_new = load_and_prepare_new(str(new_path))

    old_group = aggregate_old(df_old)
    new_group = aggregate_new(df_new)

    matches_df, matched_titles, conflict_titles, new_only_titles = build_match_sets(
        old_group, new_group
    )

    print(f"旧表标准化书名种类数: {old_group['norm_title'].nunique()}")
    print(f"新表标准化书名种类数: {new_group['norm_title'].nunique()}")
    print(f"匹配成功的书名数: {len(matched_titles)}")
    print(f"数量不一致（冲突）书名数: {len(conflict_titles)}")
    print(f"仅出现在新表的书名数: {len(new_only_titles)}")

    # 构建新表信息映射并生成匹配明细
    mapping_df = build_newinfo_mapping(df_new, matches_df)
    matched_details = apply_replacements(
        df_old=df_old,
        mapping_df=mapping_df,
        matched_titles=matched_titles,
        original_columns=original_columns,
    )

    # 生成未匹配表（包含数量冲突和仅新表）
    unmatched_items = build_unmatched_table(
        df_new=df_new,
        matches_df=matches_df,
        new_only_titles=new_only_titles,
    )

    # 生成匹配原始表：保留旧表原始字段，只附加匹配元数据，方便对照统计
    matched_original = df_old.copy()
    matched_original["_old_order"] = matched_original.index

    matched_info = matches_df.loc[
        matches_df["status"] == "matched",
        [
            "old_norm_title",
            "new_norm_title",
            "similarity",
            "new_total_quantity",
        ],
    ].rename(
        columns={
            "old_norm_title": "norm_title",
            "new_norm_title": "匹配的新表标准化书名",
            "similarity": "匹配相似度",
            "new_total_quantity": "新表数量汇总",
        }
    )

    matched_original = matched_original.merge(
        matched_info,
        on="norm_title",
        how="inner",
    )

    if "_old_order" not in matched_original.columns:
        matched_original["_old_order"] = matched_original.index

    # 引入新表顺序用于排序
    order_info = mapping_df[["norm_title", "_new_order"]].drop_duplicates(
        subset=["norm_title"]
    )
    matched_original = matched_original.merge(
        order_info,
        on="norm_title",
        how="left",
    )

    if "_new_order" in matched_original.columns and "_old_order" in matched_original.columns:
        matched_original = matched_original.sort_values(
            by=["_new_order", "_old_order"]
        )

    for helper_col in ["_new_order", "_old_order"]:
        if helper_col in matched_original.columns:
            matched_original.drop(columns=[helper_col], inplace=True)

    # 写出结果文件
    matched_details.to_excel(matched_path, index=False)
    unmatched_items.to_excel(unmatched_path, index=False)
    matched_original.to_excel(matched_original_path, index=False)

    # 写日志
    write_log(
        old_path=str(old_path),
        new_path=str(new_path),
        df_old=df_old,
        df_new=df_new,
        matched_titles=matched_titles,
        conflict_titles=conflict_titles,
        new_only_titles=new_only_titles,
        matched_path=str(matched_path),
        unmatched_path=str(unmatched_path),
        matched_original_path=str(matched_original_path),
    )

    return {
        "matched_details": matched_path,
        "unmatched_items": unmatched_path,
        "matched_original": matched_original_path,
    }


def main() -> None:
    print("开始处理 Excel 匹配任务...")
    print(f"旧表路径: {OLD_FILE_PATH}")
    print(f"新表路径: {NEW_FILE_PATH}")

    try:
        run_matching(
            old_path=OLD_FILE_PATH,
            new_path=NEW_FILE_PATH,
            matched_path=OUTPUT_MATCHED_PATH,
            unmatched_path=OUTPUT_UNMATCHED_PATH,
            matched_original_path=OUTPUT_MATCHED_ORIGINAL_PATH,
        )
    except FileNotFoundError as e:
        print(f"读取文件失败，请检查路径是否正确: {e}")
        return
    except ValueError as e:
        print(f"数据结构错误: {e}")
        return
    except Exception as e:  # noqa: BLE001
        print(f"处理过程中出错: {e}")
        return

    print("处理完成。")


if __name__ == "__main__":
    main()

