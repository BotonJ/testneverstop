
def match_subject_name(subject, alias_map):
    """
    返回包含主科目名和所有别名的候选集（用于匹配源数据中的科目行）
    :param subject: 模板中出现的科目名称
    :param alias_map: mapping_loader 提供的 subject_alias_map 字典
    :return: set[str] 所有可匹配的名称
    """
    subject = subject.strip() if isinstance(subject, str) else ""
    candidates = {subject}
    for std, aliases in alias_map.items():
        if subject == std or subject in aliases:
            candidates.add(std)
            candidates.update(aliases)
            break  # 找到一个匹配组即返回，避免混乱匹配
    return candidates
