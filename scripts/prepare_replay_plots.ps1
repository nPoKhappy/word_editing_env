#requires -Version 7.0
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$run=Join-Path $root 'outputs/chat-replay-20261028'
$research=Join-Path $run 'inputs/research'
$original='C:\Users\Administrator\Desktop\digital_twin_revised'
$paths=@('results/Tabular_MLP_New/tabular_predictions.csv','results/ann_teacher_step_validation_in_range_20261013/gain_comparison_detail.csv')
foreach($i in 19..26){$paths+="data/Claus_dynamic/Test_dataform_change_air2_R=5-$i.csv"}
foreach($p in $paths){$out=Join-Path $research $p; New-Item -ItemType Directory -Force (Split-Path $out -Parent)|Out-Null;Copy-Item -LiteralPath (Join-Path $original $p) -Destination $out}
$vis=Join-Path $research 'results/visualization/gradnorm_ratio0p005_0p01_0p015_0p02_seed42'
New-Item -ItemType Directory -Force $vis|Out-Null
Get-ChildItem (Join-Path $run 'inputs/data') -Filter 'seven_model*.csv' | Copy-Item -Destination $vis
$scripts=Join-Path $run 'plots/scripts';New-Item -ItemType Directory -Force $scripts|Out-Null
$first=Get-Content -LiteralPath (Join-Path $run 'evidence/original-scripts/plot_jtice_rmse_parity_near_untitled.py') -Raw
$first=$first.Replace("BASE = Path(r'C:\Users\Administrator\Desktop\digital_twin_revised')","BASE = Path(r'$research')")
$first|Set-Content (Join-Path $scripts 'plot_five_rmse_parity_near.py') -Encoding utf8
$out=Join-Path $run 'plots/four-model';New-Item -ItemType Directory -Force $out|Out-Null
$second=Get-Content -LiteralPath (Join-Path $run 'inputs/assets/v14_gain/plot_four_model_test_rmse.py') -Raw
$second=$second.Replace("ROOT = Path(r'C:\Users\Administrator\Desktop\digital_twin_revised')","ROOT = Path(r'$research')")
$second=$second.Replace("OUT = Path(r'D:\word_editing_env\thesis_outline\jtice_paper\assets\v14_gain')","OUT = Path(r'$out')")
$source=Join-Path $scripts 'plot_five_rmse_parity_near.py'
$second=$second.Replace("SOURCE = Path(r'D:\word_editing_env\.tmp\plot_jtice_rmse_parity_near_untitled.py')","SOURCE = Path(r'$source')")
$second|Set-Content (Join-Path $scripts 'plot_four_model_test_rmse.py') -Encoding utf8
Write-Output "Scientific plot inputs and adapted scripts: $scripts"
