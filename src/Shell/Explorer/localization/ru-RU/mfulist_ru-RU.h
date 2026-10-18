/*
 *  mfulist.h - The default MFU lists
 *
 *  The MFU lists need to be replicated for MUI purposes, so we centralize
 *  them here.
 *
 */

#define MFU_SETDEFAULTS "%ALLUSERSPROFILE%\\Главное меню\\Выбор программ по умолчанию.lnk"

//
//  32-bit Client for all user types
//
#define MFU_PRO32ALL_00 "%USERPROFILE%\\Главное меню\\Программы\\Internet Explorer.lnk"
#define MFU_PRO32ALL_01 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Стандартные\\Медиа-центр\\Медиа-центр.lnk"
#define MFU_PRO32ALL_02 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Журнал Windows.lnk"
#define MFU_PRO32ALL_03 "%ALLUSERSPROFILE%\\Главное меню\\Выбор программ по умолчанию.lnk"
#define MFU_PRO32ALL_04 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Знакомство с Tablet PC.lnk"
#define MFU_PRO32ALL_05 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Подключение к Интернету через MSN.lnk"
#define MFU_PRO32ALL_06 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\MSN Explorer.lnk"
#define MFU_PRO32ALL_07 "%USERPROFILE%\\Главное меню\\Программы\\Проигрыватель Windows Media.lnk"
#define MFU_PRO32ALL_08 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Windows Messenger.lnk"
#define MFU_PRO32ALL_09 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Знакомство с Windows XP.lnk"
#define MFU_PRO32ALL_10 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Стандартные\\Windows Movie Maker.lnk"
#define MFU_PRO32ALL_11 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Стандартные\\Служебные\\Мастер переноса файлов и параметров.lnk"
#define MFU_PRO32ALL_12 ""
#define MFU_PRO32ALL_13 ""
#define MFU_PRO32ALL_14 ""
#define MFU_PRO32ALL_15 ""

//
//  64-bit Client for all user types
//
#define MFU_PRO64ALL_00 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Стандартные\\Медиа-центр\\Медиа-центр.lnk"
#define MFU_PRO64ALL_01 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Журнал Windows.lnk"
#define MFU_PRO64ALL_02 "%ALLUSERSPROFILE%\\Главное меню\\Выбор программ по умолчанию.lnk"
#define MFU_PRO64ALL_03 "%ALLUSERSPROFILE%\\Главное меню\\Программы\\Знакомство с Tablet PC.lnk"
#define MFU_PRO64ALL_04 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Командная строка.lnk"
#define MFU_PRO64ALL_05 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Блокнот.lnk"
#define MFU_PRO64ALL_06 ""
#define MFU_PRO64ALL_07 ""
#define MFU_PRO64ALL_08 ""
#define MFU_PRO64ALL_09 ""
#define MFU_PRO64ALL_10 ""
#define MFU_PRO64ALL_11 ""
#define MFU_PRO64ALL_12 ""
#define MFU_PRO64ALL_13 ""
#define MFU_PRO64ALL_14 ""
#define MFU_PRO64ALL_15 ""

//
//  32-bit Server for administrators
//
#define MFU_SRV32ADM_00 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Командная строка.lnk"
#define MFU_SRV32ADM_01 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Блокнот.lnk"
#define MFU_SRV32ADM_02 ""
#define MFU_SRV32ADM_03 ""
#define MFU_SRV32ADM_04 ""
#define MFU_SRV32ADM_05 ""
#define MFU_SRV32ADM_06 ""
#define MFU_SRV32ADM_07 ""
#define MFU_SRV32ADM_08 ""
#define MFU_SRV32ADM_09 ""
#define MFU_SRV32ADM_10 ""
#define MFU_SRV32ADM_11 ""
#define MFU_SRV32ADM_12 ""
#define MFU_SRV32ADM_13 ""
#define MFU_SRV32ADM_14 ""
#define MFU_SRV32ADM_15 ""

//
//  64-bit Server for administrators
//
#define MFU_SRV64ADM_00 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Командная строка.lnk"
#define MFU_SRV64ADM_01 "%USERPROFILE%\\Главное меню\\Программы\\Стандартные\\Блокнот.lnk"
#define MFU_SRV64ADM_02 ""
#define MFU_SRV64ADM_03 ""
#define MFU_SRV64ADM_04 ""
#define MFU_SRV64ADM_05 ""
#define MFU_SRV64ADM_06 ""
#define MFU_SRV64ADM_07 ""
#define MFU_SRV64ADM_08 ""
#define MFU_SRV64ADM_09 ""
#define MFU_SRV64ADM_10 ""
#define MFU_SRV64ADM_11 ""
#define MFU_SRV64ADM_12 ""
#define MFU_SRV64ADM_13 ""
#define MFU_SRV64ADM_14 ""
#define MFU_SRV64ADM_15 ""

//
//  Macros
//
#define MFU_ENUM(fn, type)             \
    fn(type##_00, MFU_##type##_00)     \
    fn(type##_01, MFU_##type##_01)     \
    fn(type##_02, MFU_##type##_02)     \
    fn(type##_03, MFU_##type##_03)     \
    fn(type##_04, MFU_##type##_04)     \
    fn(type##_05, MFU_##type##_05)     \
    fn(type##_06, MFU_##type##_06)     \
    fn(type##_07, MFU_##type##_07)     \
    fn(type##_08, MFU_##type##_08)     \
    fn(type##_09, MFU_##type##_09)     \
    fn(type##_10, MFU_##type##_10)     \
    fn(type##_11, MFU_##type##_11)     \
    fn(type##_12, MFU_##type##_12)     \
    fn(type##_13, MFU_##type##_13)     \
    fn(type##_14, MFU_##type##_14)     \
    fn(type##_15, MFU_##type##_15)     \

#define MFUENUM_CSTR(nm, val)   TEXT(val),
#define MFUENUM_RCSTR(nm, val)  IDS_MFU_##nm val

#define MFU_ENUMRC(type)        MFU_ENUM(MFUENUM_RCSTR, type)
#define MFU_ENUMC(type)         MFU_ENUM(MFUENUM_CSTR,  type)
