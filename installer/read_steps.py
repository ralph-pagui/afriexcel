import json

log_path = r'C:\Users\RALPH 2 .0\.gemini\antigravity\brain\019eda48-76ef-487c-92cf-48194de34e48\.system_generated\logs\transcript.jsonl'

with open(log_path, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            data = json.loads(line)
            idx = data.get('step_index')
            if 6230 <= idx <= 6630:
                src = data.get('source')
                tp = data.get('type')
                content = data.get('content', '')
                if tp in ['USER_INPUT', 'PLANNER_RESPONSE'] and content:
                    print(f"[{idx}] {src} ({tp}): {content[:300]}\n---")
        except Exception as e:
            pass
