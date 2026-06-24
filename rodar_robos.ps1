$pasta  = "c:\Users\robson.noberto\Desktop\Controle Kips Motoristas"
$log    = "$pasta\log_agendamento.txt"
$python = (Get-Command py -ErrorAction SilentlyContinue).Source

if (-not $python) {
    Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') [ERRO] Python (py) nao encontrado."
    exit 1
}

Add-Content $log ""
Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ===== INICIO ====="

Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') Rodando Robo 1..."
& $python "$pasta\Robo controle brdrive1.py" >> $log 2>&1
Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') Robo 1 concluido."

Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') Rodando Robo 2..."
& $python "$pasta\Robo dashboard brdrive2.py" >> $log 2>&1
Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') Robo 2 concluido."

Add-Content $log "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ===== FIM ====="
