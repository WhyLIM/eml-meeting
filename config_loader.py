import os
import yaml
from typing import Dict, Any, Optional


class ConfigLoader:

    def __init__(self, config_path: str = "config.yaml"):
        self.config_path = config_path
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")

        with open(self.config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)

        return config if config else {}

    def get(self, key: str, default: Any = None) -> Any:
        return self.config.get(key, default)

    def get_input_dir(self) -> Optional[str]:
        return self.config.get('input_dir')

    def get_output_file(self) -> Optional[str]:
        return self.config.get('output_file')

    def get_api_provider(self) -> Optional[str]:
        return self.config.get('api_provider')

    def get_model(self) -> Optional[str]:
        return self.config.get('model')

    def reload(self):
        self.config = self.load_config()

    def validate(self) -> bool:
        required_keys = ['input_dir', 'output_file', 'api_provider']
        for key in required_keys:
            if key not in self.config or not self.config[key]:
                print(f"警告: 配置文件中缺少必要参数: {key}")
                return False
        return True


def load_config_from_file(config_path: str = "config.yaml") -> ConfigLoader:
    return ConfigLoader(config_path)
