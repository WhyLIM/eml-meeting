import json
import traceback
import time
from typing import Dict, Optional, List
from concurrent.futures import ThreadPoolExecutor, as_completed
from llm_client import LLMClient
from config import get_max_concurrency, MAX_RETRIES

SYSTEM_PROMPT = """你是一个专业的学术报告信息提取助手。请从邮件内容中准确提取所有学术报告或培训的信息。

重要：请提取邮件中的**所有**讲座信息，返回一个JSON数组。如果邮件中有4个讲座，就返回4个对象；如果有1个讲座，就返回1个对象。

每个讲座对象必须严格按照以下字段：
- training_name: 培训/会议名称（直接使用邮件中的主题标题）
- start_time: 开始时间，格式为"yyyy-MM-dd hh:mm"，如2025-12-25 10:00（注意：用空格分隔日期和时间，不要用T，不要秒）
- end_time: 结束时间，格式为"yyyy-MM-dd hh:mm"，如2025-12-25 11:00（注意：用空格分隔日期和时间，不要用T，不要秒），如果未明确结束时间，请设置为null
- duration_hours: 学时（持续时长），以数字表示，单位为小时，如2、1.5、3等
- location: 地点，如"线上"、"F512"、"A101"、"会议室"等
- purpose: 讲座目的，根据讲座主题（training_name）推断并概括讲座的目的，50字以内（中文，不要写"学术讲座"这种通用回答）
- content: 讲座内容，根据讲座主题（training_name）推断并概括讲座的主要内容，50字以内（中文，不要与training_name重复）

注意事项：
- **必须提取所有讲座**，不要遗漏任何一个
- purpose和content需要根据讲座主题进行合理的推断和概括，不能简单地复制主题标题或使用"学术讲座"这种通用回答
- 如果结束时间无法从邮件中提取，请设置为null
- 时间格式必须严格按照"yyyy-MM-dd hh:mm"格式输出，如"2025-12-25 10:00"
- 日期和时间之间用**空格**分隔，不要用T
- 时间只包含小时和分钟，不要包含秒
- 学时必须为数字（整数或小数），不带单位
- 讲座目的和讲座内容各概述一句话，50字以内（中文）
- 输出必须为严格的JSON数组格式，不要包含任何其他文字
- 数组中的每个对象代表一个讲座

返回格式示例（必须严格遵循此格式和字段名）：
[
  {
    "training_name": "AIGC赋能ProQuest - PQDT Global全球博硕论文的高效利用",
    "start_time": "2025-12-25 10:00",
    "end_time": "2025-12-25 11:00",
    "duration_hours": 1.0,
    "location": "线上",
    "purpose": "介绍如何利用AI技术提升论文检索效率",
    "content": "讲解PQDT Global数据库的使用方法和检索技巧"
  },
  {
    "training_name": "Understanding and Characterizing Regularization",
    "start_time": "2025-12-26 14:00",
    "end_time": null,
    "duration_hours": null,
    "location": "F512会议室",
    "purpose": "深入理解正则化技术在机器学习中的作用",
    "content": "系统讲解正则化的原理、类型和应用场景"
  }
]"""


def create_extraction_prompt(email_data: Dict[str, str]) -> str:
    prompt_parts = [
        "请从以下邮件内容中提取学术报告信息：\n\n",
    ]

    if email_data.get('subject'):
        prompt_parts.append(f"邮件主题：{email_data['subject']}\n")

    if email_data.get('from'):
        prompt_parts.append(f"发件人：{email_data['from']}\n")

    if email_data.get('date'):
        prompt_parts.append(f"邮件日期：{email_data['date']}\n")

    if email_data.get('body'):
        prompt_parts.append(f"\n邮件正文：\n{email_data['body']}\n")

    prompt_parts.append(
        "\n**重要**：请提取邮件中的**所有**讲座信息，不要遗漏任何一个。返回JSON数组，每个对象代表一个讲座。必须使用以下字段名：training_name, start_time, end_time, duration_hours, location, purpose, content。时间格式必须是\"yyyy-MM-dd hh:mm\"（空格分隔）。如果邮件中未明确结束时间，end_time必须设置为null。purpose和content需要根据讲座主题进行合理的推断和概括，不要简单地复制标题或使用\"学术讲座\"这种通用回答。只返回JSON数组，不要包含其他文字。"
    )

    return "".join(prompt_parts)


def extract_training_info(email_data: Dict[str, str],
                          api_name: str = "zai-plan") -> List[Dict[str, str]]:
    client = LLMClient(api_name)

    prompt = create_extraction_prompt(email_data)

    messages = [{"role": "user", "content": prompt}]

    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = client.chat(messages, system=SYSTEM_PROMPT)

            lectures = client.extract_json(response)

            results = []
            if isinstance(lectures, list):
                for lecture in lectures:
                    results.append({
                        'training_name':
                        lecture.get('training_name'),
                        'start_time':
                        lecture.get('start_time'),
                        'end_time':
                        lecture.get('end_time'),
                        'duration_hours':
                        lecture.get('duration_hours'),
                        'location':
                        lecture.get('location'),
                        'purpose':
                        lecture.get('purpose'),
                        'content':
                        lecture.get('content'),
                        'raw_response':
                        response
                    })
            elif isinstance(lectures, dict):
                results.append({
                    'training_name': lectures.get('training_name'),
                    'start_time': lectures.get('start_time'),
                    'end_time': lectures.get('end_time'),
                    'duration_hours': lectures.get('duration_hours'),
                    'location': lectures.get('location'),
                    'purpose': lectures.get('purpose'),
                    'content': lectures.get('content'),
                    'raw_response': response
                })
            else:
                return [{
                    'training_name': None,
                    'start_time': None,
                    'end_time': None,
                    'duration_hours': None,
                    'location': None,
                    'purpose': None,
                    'content': None,
                    'error': '无法解析JSON响应',
                    'raw_response': response
                }]

            return results

        except Exception as e:
            last_error = e
            error_type = type(e).__name__

            if error_type == 'HTTPError' and hasattr(
                    e, 'response') and e.response.status_code == 429:
                if attempt < MAX_RETRIES - 1:
                    wait_time = (2**attempt) * 5
                    print(
                        f"  429错误，等待{wait_time}秒后重试 ({attempt + 1}/{MAX_RETRIES})..."
                    )
                    time.sleep(wait_time)
                else:
                    print(f"  429错误，已达最大重试次数")
            elif attempt < MAX_RETRIES - 1:
                wait_time = 2**attempt
                print(
                    f"  {error_type}，等待{wait_time}秒后重试 ({attempt + 1}/{MAX_RETRIES})..."
                )
                time.sleep(wait_time)

    return [{
        'training_name':
        None,
        'start_time':
        None,
        'end_time':
        None,
        'duration_hours':
        None,
        'location':
        None,
        'purpose':
        None,
        'content':
        None,
        'error':
        f"{type(last_error).__name__}: {str(last_error)}"
        if last_error else "未知错误",
        'traceback':
        traceback.format_exc() if last_error else ""
    }]


def _extract_single(email_data: Dict[str, str], api_name: str, index: int,
                    total: int) -> tuple[int, List[Dict]]:
    try:
        lectures = extract_training_info(email_data, api_name)
        for lecture in lectures:
            lecture['file_path'] = email_data.get('file_path', '')
            lecture['file_name'] = email_data.get('file_name', '')
        return index, lectures
    except Exception as e:
        error_result = [{
            'training_name': None,
            'start_time': None,
            'end_time': None,
            'duration_hours': None,
            'location': None,
            'purpose': None,
            'content': None,
            'file_path': email_data.get('file_path', ''),
            'file_name': email_data.get('file_name', ''),
            'error': f"{type(e).__name__}: {str(e)}",
            'traceback': traceback.format_exc()
        }]
        return index, error_result


def extract_training_info_batch(email_data_list: list,
                                api_name: str = "zai-plan",
                                progress_callback=None) -> list:
    max_concurrency = get_max_concurrency(api_name)
    total = len(email_data_list)

    results = [None] * total

    with ThreadPoolExecutor(max_workers=max_concurrency) as executor:
        future_to_index = {
            executor.submit(_extract_single, email_data, api_name, i, total): i
            for i, email_data in enumerate(email_data_list)
        }

        for future in as_completed(future_to_index):
            index, lectures = future.result()
            results[index] = lectures

            if progress_callback:
                completed = sum(1 for r in results if r is not None)
                progress_callback(completed, total,
                                  email_data_list[index].get('file_name', ''))

    final_results = []
    for r in results:
        if r is not None:
            final_results.extend(r)

    return final_results
