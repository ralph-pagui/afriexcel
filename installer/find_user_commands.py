import json

log_path = r'C:\Users\RALPH 2 .0\.gemini\antigravity\brain\019eda48-76ef-487c-92cf-48194de34e48\.system_generated\logs\transcript.jsonl'

with open(log_path, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get('source') == 'USER_EXPLICIT' and data.get('type') == 'USER_INPUT':
                content = data.get('content', '')
                if any(w in content.lower() for w in ['commande', 'install', 'ancien', 'marche', 'dossier', 'addin', 'excel']):
                    print(f"Step {data.get('step_index')}: {content}")
        except Exception as e:
            pass
