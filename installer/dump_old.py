import subprocess

out = subprocess.check_output(['git', 'show', '743557f:addin-frontend/public/install.ps1'], encoding='utf-8', errors='ignore')
print(out[:3000])
with open(r'F:\cop\installer\old_install_743557f.ps1', 'w', encoding='utf-8') as f:
    f.write(out)
print("Saved old install script to F:\\cop\\installer\\old_install_743557f.ps1")
