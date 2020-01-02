# CFTRDataTransfer
Transfer of CFTR74 v3 on Agena MassArray from Agena Typer to Cerner Millennium Helix

1. Batch created in Millennium

2. Import script run from worksheet which prompts user to select the file to import from
- MALDITOF_CFTR_Analysis_Import_Bootstrap.vbs

3. Import script loads the above bootstrap which utilises the following config file and the scriptwise file which includes a number of functions
- MALDITOF_CFTR_Analysis_Import_Config / ScriptWise.vbs

4. Config file loads the following vbs script
- MALDITOF_CFTR_Analysis_Import.vbs

5. The above includes the following vbs scripts which include various subprocesses
- MALDITOF_CFTR_Common_Import.vbs
- MALDITOF_CFTR_PolyT_Import.vbs
- HelixLib.vbs
