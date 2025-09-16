def format_fio(full_name):
    """Форматирование ФИО в формате 'Фамилия И.О.'"""
    parts = full_name.split(' ')
    last_name = parts[0]
    initials = ''.join(f"{name[0]}." for name in parts[1:])
    return f"{last_name} {initials}"