"""保存模板、主数据、汇总主表、输出目录及最近供应商，兼容旧配置。"""
import json
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass
class Settings:
    excel1_path: str = ""
    excel2_path: str = ""
    last_supplier: str = "NEXPERIA"
    master_data_path: str = ""
    output_root: str = ""
    supplier_image_dir: str = ""


def load_settings(path: Path) -> Settings:
    if not path.exists():
        return Settings()
    data = json.loads(path.read_text(encoding="utf-8"))
    required = {"excel1_path", "excel2_path", "last_supplier"}
    optional = {"master_data_path", "output_root", "supplier_image_dir"}
    if not isinstance(data, dict) or not required <= set(data) or set(data) - (required | optional):
        raise ValueError("配置缺少必填字段或含有未知字段")
    if any(not isinstance(value, str) for value in data.values()):
        raise ValueError("配置字段必须为字符串")
    return Settings(**data)


def save_settings(path: Path, settings: Settings) -> None:
    data = asdict(settings)
    if any(not isinstance(value, str) for value in data.values()):
        raise ValueError("配置字段必须为字符串")
    path.parent.mkdir(parents=True, exist_ok=True)
    # 同目录替换，避免中断留下半份配置。
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)
