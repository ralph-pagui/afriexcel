import json

log_path = r'C:\Users\RALPH 2 .0\.gemini\antigravity\brain\019eda48-76ef-487c-92cf-48194de34e48\.system_generated\logs\transcript.jsonl'

with open(log_path, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            data = json.loads(line)
            content = data.get('content', '')
            idx = data.get('step_index')
            src = data.get('source')
            if src == 'USER_EXPLICIT' and any(k in content.lower() for k in ['ça marche', 'marche bien', 'fonctionne', 'super', 'parfait']):
                print(f"[{idx}] {content[:200]}")
        except:
            pass
