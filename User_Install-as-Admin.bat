@echo off
rem �J�����g�f�B���N�g�����o�b�`�t�@�C���Ɠ����ꏊ�ɂ���C�f�B�I���B
rem �Ǘ��Ҍ�������Ńo�b�`�t�@�C�����N������ƁA�J�����g�f�B���N�g�����ʂ̏ꏊ�Ɉړ����Ă��܂��̂ŁA���̏������K�v�ƂȂ�B
cd /d %~dp0
powershell -NoLogo -ExecutionPolicy RemoteSigned -File .\Setup.ps1 -ActionType Install
pause
