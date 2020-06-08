#
import os
import sys

from shutil import copyfile


vestas_dll_dir = r"C:\Users\didng\OneDrive - Vestas Wind Systems A S\Projects\01_Lib\Vestas_dlls"
wtg_source_code_path = r"C:\Users\didng\OneDrive - Vestas Wind Systems A S\Projects\Source Code\01. WTG"
ppc_source_code_path = r"C:\Users\didng\OneDrive - Vestas Wind Systems A S\Projects\Source Code\02. PPC"
fbc_user_code_path = r"C:\Users\didng\OneDrive - Vestas Wind Systems A S\Projects\Source Code\03. " \
                     r"Miscellaneous\FBC User Code"
#
sys_path_psse = r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_psse)
import psse_env_manager

#
cwd = os.getcwd()

#
user_code = r"""C:\Users\didng\OneDrive - Vestas Wind Systems A S\Projects\Lib\User_code"""
sys.path.append(user_code)
import usercode27 as uc


def vestas_ppc_compilation(ppc_legacy, ppc_version, psse_ver_major, psse_ver_minor, ppc_compiler):
    vestas_ppc_dll_name = ppc_version + str(psse_ver_major) + "." + str(psse_ver_minor) + ".dll"
    vestas_ppc_dll_file = os.path.join(vestas_dll_dir, vestas_ppc_dll_name)
    #
    ppc_working_dir = os.path.join(ppc_source_code_path, ppc_version)
    if abs(ppc_legacy - 1.0) < 0.5:
        for i in os.listdir(ppc_working_dir):
            base_file, ext = os.path.splitext(i)
            if ext == ".for":
                files = os.path.join(ppc_working_dir, i)
                flx_file = os.path.join(ppc_working_dir, base_file + '.flx')
                copyfile(files, flx_file)
        flx_files = uc.get_files_name(ppc_working_dir, ['.flx'])
    #
    if abs(ppc_compiler - 1) < 0.1:
        try:
            os.remove(vestas_ppc_dll_file)
        except WindowsError:
            print 'no file *.dll to delete!', vestas_ppc_dll_name
        ierr = psse_env_manager.create_dll(34, flx_files,
                                           modsources=[],
                                           objlibfiles=[],
                                           dllname=vestas_ppc_dll_file, workdir=os.getcwd(), showprg=False,
                                           useivfvrsn='latest',
                                           shortname=vestas_ppc_dll_file, description='User Model', majorversion=1,
                                           minorversion=0, buildversion=0, companyname='Vestas', mypathlib=False)
        uc.dirCreateClean(vestas_dll_dir, ["*.f"])
    else:
        print "compilation is not required"


def vestas_wtg_compilation(wtg_version, psse_ver_major, psse_ver_minor, wtg_compiler):
    vestas_wtg_dll_name = wtg_version + str(psse_ver_major) + "." + str(psse_ver_minor) + ".dll"
    vestas_wtg_dll_file = os.path.join(vestas_dll_dir, vestas_wtg_dll_name)

    #
    wtg_working_dir = os.path.join(wtg_source_code_path, wtg_version)
    fortran_files = uc.get_files_name(wtg_working_dir, ['.for'])
    if abs(wtg_compiler - 1) < 0.1:
        try:
            os.remove(vestas_wtg_dll_file)
        except WindowsError:
            print 'no file *.dll to delete!', vestas_wtg_dll_name
        ierr = psse_env_manager.create_dll(34, fortran_files,
                                           modsources=[],
                                           objlibfiles=[],
                                           dllname=vestas_wtg_dll_file, workdir=os.getcwd(), showprg=False,
                                           useivfvrsn='latest',
                                           shortname=vestas_wtg_dll_file, description='User Model', majorversion=1,
                                           minorversion=0, buildversion=0, companyname='Vestas', mypathlib=False)
        uc.dirCreateClean(vestas_dll_dir, ["*.f"])
    else:
        print "compilation is not required"


def vestas_msufbk_compilation(msufbc_version, psse_ver_major, psse_ver_minor, msufbc_compiler):
    vestas_msufbc_dll_name = msufbc_version + str(psse_ver_major) + "." + str(psse_ver_minor) + ".dll"
    vestas_msufbc_version_dll_file = os.path.join(vestas_dll_dir, vestas_msufbc_dll_name)

    #
    wtg_working_dir = os.path.join(fbc_user_code_path, msufbc_version)
    fortran_files = uc.get_files_name(wtg_working_dir, ['.for'])

    if abs(msufbc_compiler - 1) < 0.1:
        try:
            os.remove(vestas_msufbc_version_dll_file)
        except WindowsError:
            print 'no file *.dll to delete!', vestas_msufbc_version_dll_file
        ierr = psse_env_manager.create_dll(34, fortran_files,
                                           modsources=[],
                                           objlibfiles=[],
                                           dllname=vestas_msufbc_version_dll_file, workdir=os.getcwd(), showprg=False,
                                           useivfvrsn='latest',
                                           shortname=vestas_msufbc_version_dll_file, description='User Model',
                                           majorversion=1,
                                           minorversion=0, buildversion=0, companyname='Vestas', mypathlib=False)
        uc.dirCreateClean(vestas_dll_dir, ["*.f"])
    else:
        print "compilation is not required"
