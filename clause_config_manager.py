# -*- coding: utf-8 -*-
"""
配置管理模块 - 支持外部 JSON 配置文件
支持：加载、保存、热更新、默认值回退

Author: Dachi Yijin
Date: 2025-12-21
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field, asdict
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


@dataclass
class MatchThresholds:
    """匹配阈值配置"""
    exact_min: float = 0.98
    semantic_min: float = 0.85
    keyword_min: float = 0.60
    fuzzy_min: float = 0.40
    accept_min: float = 0.15


class ClauseConfigManager:
    """
    配置管理器 - 单例模式

    使用方式:
        config = ClauseConfigManager.get_instance()
        config.load()  # 加载配置

        # 获取映射
        cn_name = config.get_client_mapping("reinstatement value")

        # 添加新映射
        config.add_client_mapping("new term", "新术语")
        config.save()  # 保存到文件
    """

    _instance: Optional['ClauseConfigManager'] = None

    # ========================================
    # 默认配置（JSON 不存在时的回退值）
    # ========================================
    DEFAULT_CLIENT_EN_CN_MAP: Dict[str, str] = {
        "interpretation & headings": "通译和标题条款",
        "reinstatement value": "重置价值条款",
        "reinstatement value clause": "重置价值条款",
        "replacement value": "重置价值条款",
        "time adjustment": "72小时条款",
        "time adjustment (72 hours)": "72小时条款",
        "72 hours clause": "72小时条款",
        "civil authorities clause": "公共当局扩展条款",
        "civil authorities": "公共当局扩展条款",
        "public authorities clause": "公共当局扩展条款",
        "errors and omissions clause": "错误和遗漏条款",
        "errors and omissions": "错误和遗漏条款",
        "loss notification clause": "损失通知条款",
        "loss notification": "损失通知条款",
        "no control clause": "不受控制条款",
        "no control": "不受控制条款",
        "60 days' notice of cancellation by insurer": "60天通知注销保单条款",
        "expediting costs": "加快费用条款",
        "all other contents": "其它物品条款",
        "alterations, additions and repairs": "变更和维修条款",
        "escalation": "自动升值扩展条款",
        "automatic reinstatement of sum insured": "自动恢复保险金额条款",
        "removal of debris": "清理残骸费用扩展条款",
        "strike, riot, civil commotion": "罢工、暴动或民众骚乱条款",
        "srcc": "罢工、暴动或民众骚乱条款",
        "earthquake and tsunami": "地震扩展条款",
        "theft and robbery": "盗窃、抢劫扩展条款",
        "professional fees": "专业费用及索赔准备费用条款",
    }

    DEFAULT_SEMANTIC_ALIAS_MAP: Dict[str, str] = {
        "污染保险": "意外污染责任",
        "污染责任": "意外污染责任",
        "露天财产": "露天及简易建筑内存放财产",
        "损害防止": "阻止损失",
        "施救费用": "阻止损失",
        "崩塌沉降": "地面突然下陷下沉",
        "地面下陷": "地面突然下陷下沉",
        "重置(价值)": "重置价值",
        "公共当局": "公共当局扩展",
    }

    DEFAULT_KEYWORD_MAP: Dict[str, List[str]] = {
        "污染": ["污染", "意外污染", "pollution"],
        "地震": ["地震", "震动", "earthquake"],
        "海啸": ["海啸", "tsunami"],
        "盗窃": ["盗窃", "盗抢", "抢劫", "burglary", "theft", "robbery"],
        "洪水": ["洪水", "水灾", "flood"],
        "火灾": ["火灾", "火险", "fire"],
        "重置": ["重置", "重建", "reinstatement", "replacement"],
    }

    DEFAULT_PENALTY_KEYWORDS: List[str] = ["打孔盗气"]

    DEFAULT_NOISE_WORDS: List[str] = [
        "企业财产保险", "附加", "扩展", "条款", "险",
        "（A款）", "（B款）", "(A款)", "(B款)",
        "2025版", "2024版", "2023版", "版",
        "clause", "extension", "cover", "insurance",
    ]

    def __init__(self):
        """初始化配置管理器"""
        self._config_path: Optional[Path] = None
        self._is_loaded: bool = False

        # 运行时配置（会合并默认值和外部配置）
        self.thresholds = MatchThresholds()
        self.client_en_cn_map: Dict[str, str] = {}
        self.semantic_alias_map: Dict[str, str] = {}
        self.keyword_extract_map: Dict[str, List[str]] = {}
        self.exact_clause_map: Dict[str, str] = {}
        self.penalty_keywords: List[str] = []
        self.noise_words: List[str] = []

        # 记录用户新增的映射（用于保存）
        self._user_additions: Dict[str, Dict] = {
            'client_en_cn_map': {},
            'semantic_alias_map': {},
            'exact_clause_map': {},
        }

    @classmethod
    def get_instance(cls) -> 'ClauseConfigManager':
        """获取单例实例"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @classmethod
    def reset_instance(cls):
        """重置单例（用于测试）"""
        cls._instance = None

    def _get_default_config_path(self) -> Path:
        """获取默认配置文件路径"""
        # 优先查找与脚本同目录的配置
        script_dir = Path(__file__).parent
        config_path = script_dir / "clause_config.json"

        if config_path.exists():
            return config_path

        # 其次查找用户目录
        user_config = Path.home() / ".clause_diff" / "config.json"
        if user_config.exists():
            return user_config

        # 返回默认位置（可能不存在）
        return config_path

    def load(self, config_path: Optional[str] = None) -> bool:
        """
        加载配置文件

        Args:
            config_path: 配置文件路径，None 则使用默认路径

        Returns:
            是否成功加载外部配置
        """
        # 先加载默认值
        self._load_defaults()

        # 确定配置文件路径
        if config_path:
            self._config_path = Path(config_path)
        else:
            self._config_path = self._get_default_config_path()

        # 尝试加载外部配置
        if self._config_path.exists():
            try:
                with open(self._config_path, 'r', encoding='utf-8') as f:
                    external_config = json.load(f)

                self._merge_config(external_config)
                self._is_loaded = True
                logger.info(f"已加载配置文件: {self._config_path}")
                return True

            except json.JSONDecodeError as e:
                logger.error(f"配置文件 JSON 格式错误: {e}")
            except Exception as e:
                logger.error(f"加载配置文件失败: {e}")
        else:
            logger.info(f"配置文件不存在，使用默认配置: {self._config_path}")

        self._is_loaded = True
        return False

    def _load_defaults(self):
        """加载默认配置"""
        self.client_en_cn_map = self.DEFAULT_CLIENT_EN_CN_MAP.copy()
        self.semantic_alias_map = self.DEFAULT_SEMANTIC_ALIAS_MAP.copy()
        self.keyword_extract_map = {k: v.copy() for k, v in self.DEFAULT_KEYWORD_MAP.items()}
        self.penalty_keywords = self.DEFAULT_PENALTY_KEYWORDS.copy()
        self.noise_words = self.DEFAULT_NOISE_WORDS.copy()
        self.exact_clause_map = {}
        self.thresholds = MatchThresholds()

    def _merge_config(self, external: Dict[str, Any]):
        """合并外部配置（外部配置覆盖默认值）"""

        # 合并阈值
        if 'thresholds' in external:
            thresh = external['thresholds']
            self.thresholds = MatchThresholds(
                exact_min=thresh.get('exact_min', self.thresholds.exact_min),
                semantic_min=thresh.get('semantic_min', self.thresholds.semantic_min),
                keyword_min=thresh.get('keyword_min', self.thresholds.keyword_min),
                fuzzy_min=thresh.get('fuzzy_min', self.thresholds.fuzzy_min),
                accept_min=thresh.get('accept_min', self.thresholds.accept_min),
            )

        # 合并映射字典（外部覆盖默认）
        if 'client_en_cn_map' in external:
            ext_map = {k: v for k, v in external['client_en_cn_map'].items()
                      if not k.startswith('_')}  # 忽略注释字段
            self.client_en_cn_map.update(ext_map)

        if 'semantic_alias_map' in external:
            ext_map = {k: v for k, v in external['semantic_alias_map'].items()
                      if not k.startswith('_')}
            self.semantic_alias_map.update(ext_map)

        if 'keyword_extract_map' in external:
            ext_map = {k: v for k, v in external['keyword_extract_map'].items()
                      if not k.startswith('_')}
            self.keyword_extract_map.update(ext_map)

        if 'exact_clause_map' in external:
            ext_map = {k: v for k, v in external['exact_clause_map'].items()
                      if not k.startswith('_')}
            self.exact_clause_map.update(ext_map)

        if 'penalty_keywords' in external:
            # 合并去重
            self.penalty_keywords = list(set(
                self.penalty_keywords + external['penalty_keywords']
            ))

        if 'noise_words' in external:
            self.noise_words = list(set(
                self.noise_words + external['noise_words']
            ))

    def save(self, config_path: Optional[str] = None) -> bool:
        """
        保存配置到文件

        Args:
            config_path: 保存路径，None 则使用当前加载路径

        Returns:
            是否保存成功
        """
        save_path = Path(config_path) if config_path else self._config_path

        if not save_path:
            save_path = self._get_default_config_path()

        # 确保目录存在
        save_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            config_data = {
                "_meta": {
                    "version": "1.0",
                    "description": "智能条款比对工具配置文件",
                    "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                },
                "thresholds": asdict(self.thresholds),
                "client_en_cn_map": self.client_en_cn_map,
                "semantic_alias_map": self.semantic_alias_map,
                "keyword_extract_map": self.keyword_extract_map,
                "exact_clause_map": self.exact_clause_map,
                "penalty_keywords": self.penalty_keywords,
                "noise_words": self.noise_words,
            }

            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)

            logger.info(f"配置已保存: {save_path}")
            return True

        except Exception as e:
            logger.error(f"保存配置失败: {e}")
            return False

    # ========================================
    # 便捷访问方法
    # ========================================

    def get_client_mapping(self, english_term: str) -> Optional[str]:
        """获取英文术语的中文翻译"""
        normalized = english_term.lower().strip()
        return self.client_en_cn_map.get(normalized)

    def add_client_mapping(self, english_term: str, chinese_term: str):
        """添加新的英中映射"""
        normalized = english_term.lower().strip()
        self.client_en_cn_map[normalized] = chinese_term
        self._user_additions['client_en_cn_map'][normalized] = chinese_term

    def get_semantic_alias(self, term: str) -> Optional[str]:
        """获取语义别名"""
        for alias, target in self.semantic_alias_map.items():
            if alias in term:
                return target
        return None

    def add_semantic_alias(self, alias: str, target: str):
        """添加语义别名"""
        self.semantic_alias_map[alias] = target
        self._user_additions['semantic_alias_map'][alias] = target

    def get_exact_clause_mapping(self, clause_name: str) -> Optional[str]:
        """获取精确条款映射"""
        for src, tgt in self.exact_clause_map.items():
            if src in clause_name:
                return tgt
        return None

    def add_exact_clause_mapping(self, source: str, target: str):
        """添加精确条款映射"""
        self.exact_clause_map[source] = target
        self._user_additions['exact_clause_map'][source] = target

    def is_penalty_keyword(self, text: str) -> bool:
        """检查是否包含惩罚关键词"""
        return any(kw in text for kw in self.penalty_keywords)

    def get_keywords_for_text(self, text: str) -> set:
        """从文本提取关键词"""
        keywords = set()
        text_lower = text.lower()
        for core, variants in self.keyword_extract_map.items():
            for v in variants:
                if v.lower() in text_lower:
                    keywords.add(core)
                    break
        return keywords

    def clean_noise_words(self, text: str) -> str:
        """清理噪音词"""
        result = text
        for word in self.noise_words:
            result = result.replace(word, "").replace(word.lower(), "")
        return result.strip()

    # ========================================
    # 状态查询
    # ========================================

    @property
    def is_loaded(self) -> bool:
        """配置是否已加载"""
        return self._is_loaded

    @property
    def config_path(self) -> Optional[Path]:
        """当前配置文件路径"""
        return self._config_path

    @property
    def has_user_additions(self) -> bool:
        """是否有用户新增的映射"""
        return any(len(v) > 0 for v in self._user_additions.values())

    def get_stats(self) -> Dict[str, int]:
        """获取配置统计"""
        return {
            'client_mappings': len(self.client_en_cn_map),
            'semantic_aliases': len(self.semantic_alias_map),
            'keyword_rules': len(self.keyword_extract_map),
            'exact_mappings': len(self.exact_clause_map),
            'penalty_keywords': len(self.penalty_keywords),
            'noise_words': len(self.noise_words),
        }


# ========================================
# 便捷函数
# ========================================

def get_config() -> ClauseConfigManager:
    """获取配置管理器实例"""
    config = ClauseConfigManager.get_instance()
    if not config.is_loaded:
        config.load()
    return config


# ========================================
# 测试代码
# ========================================

if __name__ == '__main__':
    # 简单测试
    logging.basicConfig(level=logging.INFO)

    config = get_config()

    print("=== 配置统计 ===")
    for key, val in config.get_stats().items():
        print(f"  {key}: {val}")

    print("\n=== 测试映射查询 ===")
    test_terms = [
        "reinstatement value",
        "civil authorities clause",
        "unknown term",
    ]
    for term in test_terms:
        result = config.get_client_mapping(term)
        print(f"  '{term}' -> {result or '(未找到)'}")

    print("\n=== 测试关键词提取 ===")
    test_text = "地震及海啸导致的火灾损失"
    keywords = config.get_keywords_for_text(test_text)
    print(f"  '{test_text}' -> {keywords}")

    print("\n=== 测试添加新映射 ===")
    config.add_client_mapping("test clause", "测试条款")
    print(f"  添加后查询: {config.get_client_mapping('test clause')}")

    # 保存测试
    # config.save()
