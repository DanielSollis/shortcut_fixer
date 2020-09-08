import os
import datetime

import win32com.client
from difflib import get_close_matches
from Levenshtein import distance


def shortcuts_and_files(file: str) -> tuple:
    walk = os.walk(file)
    all_shortcuts = []
    all_files = set()
    for dirpath, dnames, fnames in walk:
        for file in fnames:
            file = dirpath + '\\' + file
            if file.endswith('.lnk'):
                all_shortcuts.append(file)
            else:
                all_files.add(file)
    return all_shortcuts, all_files


def broken_shortcuts(all_shortcuts: list) -> list:
    bad_shortcuts = []
    shell = win32com.client.Dispatch("WScript.Shell")
    for shortcut in all_shortcuts:
        try:
            source_shortcut = shell.CreateShortcut(shortcut)
            target = source_shortcut.Targetpath
            if not os.path.exists(target):
                bad_shortcuts.append(shortcut)
        except:
            print(shortcut)
    return bad_shortcuts


def find_originals(shortcuts: list, files: set, path) -> dict:
    originals = {}
    for shortcut in shortcuts:
        target = get_close_matches(shortcut, files)[0]
        target_tail = target.split('\\')[-1]
        shortcut_tail = shortcut.split('\\')[-1].split('Shortcut')[0]
        if distance(target_tail, shortcut_tail) < 30:
            shortcut_name_folder = shortcut[len(path):].split('\\')[1]
            target_name_folder = target[len(path):].split('\\')[1]
            if shortcut_name_folder != target_name_folder:
                originals[shortcut] = target
    return originals


def replace(original_mapping: dict) -> None:
    with(open('replacement_log_{}'.format(datetime.date.today()), 'w')) as f:
        shell = win32com.client.Dispatch("WScript.Shell")
        for shortcut, target in original_mapping.items():
            os.remove(shortcut)
            shortcut = shell.CreateShortCut(shortcut)
            shortcut.Targetpath = target
            shortcut.save()
            try:
                f.write('Replaced      {}\n\twith      {}\n\n'.format(shortcut, target))
            except:
                print(shortcut, target)


path = r'E:\Media\TV Shows\Xenoblade\Season 2'
shortcuts, files = shortcuts_and_files(path)
shortcuts = broken_shortcuts(shortcuts)
originals = find_originals(shortcuts, files, path)
replace(originals)
