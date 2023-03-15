import os
import numpy as np
import pandas as pd


def get_type_application(names, num_rows):
    types = []
    for name in names:
        types += [name] * num_rows
    return pd.DataFrame(types, columns=['Typ zadosti'])


def get_type_submission(name, num_rows, num):
    return pd.DataFrame([name] * num_rows * num, columns=['Typ podani'])


def read_excel(path_to_table, columns, mask, columns_names, num_rows):
    table = pd.read_excel(path_to_table, sheet_name='PnB procesy', nrows=num_rows,
                          usecols=columns, skiprows=mask, names=columns_names)
    return table


def duplicate_proces_activity(df, indices):
    # indices - [[value_index, start, end],]
    for index in indices:
        df.iloc[index[1]:index[2], :] = df.iloc[index[0], :]
    return df


def get_submission(path_to_table):
    num_types = 7
    num_rows = 41
    num_rows_submission = 34
    mask_processes = [0, 1, 9, 16, 25, 31, 35, 37, 41]

    proces_aktivity = read_excel(path_to_table, 'A, B', mask_processes, ['Proces', 'Aktivity'], num_rows)
    podproces_podaktiviy = proces_aktivity.copy()
    podproces_podaktiviy.rename(columns={'Proces': 'Podproces', 'Aktivity': 'PodAktivity'}, errors="raise",
                                inplace=True)

    mask_subprocesses = list(range(2, 6)) + list(range(7, 12)) + list(range(13, 20)) + \
        list(range(21, 25)) + list(range(26, 28)) + list(range(30, 32))
    proces_aktivity.iloc[mask_subprocesses, [0, 1]] = np.NAN
    proces_aktivity = duplicate_proces_activity(proces_aktivity, [[1, 2, 6], [6, 7, 12], [12, 13, 20],
                                                                  [20, 21, 25], [25, 26, 28], [29, 30, 32]])

    mask_subprocesses = set(range(34)) - set(mask_subprocesses)
    podproces_podaktiviy.iloc[list(mask_subprocesses), [0, 1]] = np.NAN

    roles = read_excel(path_to_table, 'C', mask_processes, ['Role'], num_rows)
    proces_aktivity_role = pd.concat([proces_aktivity, podproces_podaktiviy, roles], axis=1)

    fyzicka_zadost = read_excel(path_to_table, 'D, E', mask_processes, ['nekompletní', 'kompletní'], num_rows)
    type_submission = get_type_submission('Fyzicky', num_rows_submission, 2)
    type_application = get_type_application(['Nekompletni', 'Kompletni'], num_rows_submission)
    time = pd.concat([fyzicka_zadost.iloc[:, 0], fyzicka_zadost.iloc[:, 1]], axis=0).reset_index(drop=True)
    fyzicka_zadost = pd.concat([time, type_submission, type_application], axis=1)

    datova_zadost = read_excel(path_to_table, 'H, I', mask_processes, ['nekompletní', 'kompletní'], num_rows)
    type_submission = get_type_submission('Datova schranka', num_rows_submission, 2)
    type_application = get_type_application(['Nekompletni', 'Kompletni'], num_rows_submission)
    time = pd.concat([datova_zadost.iloc[:, 0], datova_zadost.iloc[:, 1]], axis=0).reset_index(drop=True)
    datova_zadost = pd.concat([time, type_submission, type_application], axis=1)

    robot_zadost = read_excel(path_to_table, 'K, L, M', mask_processes, ['nezadaná', 'částečně', 'úplně'], num_rows)
    type_submission = get_type_submission('Robot', num_rows_submission, 3)
    type_application = get_type_application(['Nezadaná', 'Částečně zadaná', 'Úplně zadaná'], num_rows_submission)
    time = pd.concat([robot_zadost.iloc[:, 0], robot_zadost.iloc[:, 1],
                      robot_zadost.iloc[:, 2]], axis=0).reset_index(drop=True)
    robot_zadost = pd.concat([time, type_submission, type_application], axis=1)

    applications = pd.concat([fyzicka_zadost, datova_zadost, robot_zadost])
    proces_aktivity_role = pd.concat([proces_aktivity_role] * num_types)
    proces_aktivity_role = proces_aktivity_role.reset_index(drop=True)
    applications = applications.reset_index(drop=True)
    podani_zadosti = pd.concat([proces_aktivity_role, applications], axis=1)
    podani_zadosti.rename(columns={0: 'Cas'}, inplace=True)
    nazev_procesu = pd.DataFrame(['Podání žádosti'] * num_rows_submission * num_types, columns=['Nazev procesu'])
    podani_zadosti = pd.concat([podani_zadosti, nazev_procesu], axis=1)

    return podani_zadosti


def get_evaluation(path_to_table):
    num_rows_evaluation = 7
    evaluation_application = read_excel(path_to_table, 'A:D', list(range(48)),
                                        ['Proces', 'Aktivity', 'Role', 'Cas'], num_rows_evaluation)
    nazev_procesu = pd.DataFrame(['Vyhodnocení žádosti'] * num_rows_evaluation, columns=['Nazev procesu'])
    evaluation_application = pd.concat([evaluation_application, nazev_procesu], axis=1)
    return evaluation_application


def get_changes(path_to_table):
    num_types = 7
    num_rows_changes = 9
    zmeny_zadosti = read_excel(path_to_table, 'A:C, E:F, I:J, L:N', list(range(61)), None, num_rows_changes)
    zmeny_zadosti.columns = [['Proces', 'Aktivity', 'Role', 'A', 'B', 'C', 'D', 'E', 'F', 'G']]

    type_submission = get_type_submission('Fyzicky', num_rows_changes, 2)
    type_application = get_type_application(['Nekompletni', 'Kompletni'], num_rows_changes)
    time = pd.concat([zmeny_zadosti.iloc[:, 3], zmeny_zadosti.iloc[:, 4]], axis=0).reset_index(drop=True)
    fyzicka_zadost = pd.concat([time, type_submission, type_application], axis=1)

    type_submission = get_type_submission('Datova schranka', num_rows_changes, 2)
    type_application = get_type_application(['Nekompletni', 'Kompletni'], num_rows_changes)
    time = pd.concat([zmeny_zadosti.iloc[:, 5], zmeny_zadosti.iloc[:, 6]], axis=0).reset_index(drop=True)
    datova_zadost = pd.concat([time, type_submission, type_application], axis=1)

    type_submission = get_type_submission('Robot', num_rows_changes, 3)
    type_application = get_type_application(['Nezadaná', 'Částečně zadaná', 'Úplně zadaná'], num_rows_changes)
    time = pd.concat([zmeny_zadosti.iloc[:, 7], zmeny_zadosti.iloc[:, 8],
                      zmeny_zadosti.iloc[:, 9]], axis=0).reset_index(drop=True)
    robot_zadost = pd.concat([time, type_submission, type_application], axis=1)

    applications = pd.concat([fyzicka_zadost, datova_zadost, robot_zadost]).reset_index(drop=True)
    proces_aktivity_role = zmeny_zadosti[['Proces', 'Aktivity', 'Role']]
    proces_aktivity_role.columns = ['Proces', 'Aktivity', 'Role']
    proces_aktivity_role = pd.concat([proces_aktivity_role] * num_types)
    proces_aktivity_role = proces_aktivity_role.reset_index(drop=True)
    proces_aktivity_role = pd.concat([proces_aktivity_role, applications], axis=1)
    proces_aktivity_role.rename(columns={0: 'Cas'}, inplace=True)
    nazev_procesu = pd.DataFrame(['Změny žádosti'] * num_rows_changes * num_types, columns=['Nazev procesu'])
    changes_application = pd.concat([proces_aktivity_role, nazev_procesu], axis=1)

    return changes_application


def get_other_changes(path_to_table):
    num_rows = 4
    other_changes = read_excel(path_to_table, 'A:D', list(range(75)), ['Proces', 'Aktivity', 'Role', 'Cas'], num_rows)
    nazev_procesu = pd.DataFrame(['Další změny v žádosti'] * num_rows, columns=['Nazev procesu'])
    other_changes = pd.concat([other_changes, nazev_procesu], axis=1)
    return other_changes


def get_df_from_table(region, table_name):
    path_to_table = '...//Pruzkum_data//' + region + '//' + table_name
    num_rows = 312

    podani_zadosti = get_submission(path_to_table)
    vyhodnoceni_zadosti = get_evaluation(path_to_table)
    zmeny = get_changes(path_to_table)
    dalsi_zmeny = get_other_changes(path_to_table)
    df = pd.concat([podani_zadosti, vyhodnoceni_zadosti, zmeny, dalsi_zmeny], ignore_index=True, sort=False)

    regions = pd.DataFrame([region] * num_rows, columns=['Kraj'])
    sources = pd.DataFrame([table_name] * num_rows, columns=['Zdroj'])
    df = pd.concat([df, regions, sources], axis=1)

    return df


if __name__ == '__main__':
    path = '...//Pruzkum_data'
    folders = os.listdir(path)
    result = pd.DataFrame()

    for folder in folders:
        files = os.listdir(path + '//' + folder)
        print(folder)
        for file in files:
            print(file)
            data_frame = get_df_from_table(folder, file)
            result = pd.concat([result, data_frame], ignore_index=True, sort=False)

    result[["Cas"]] = result[["Cas"]].apply(pd.to_numeric, errors='coerce')

    writer = pd.ExcelWriter('table.xlsx', engine='xlsxwriter')
    result.to_excel(writer, sheet_name='PnB procesy', startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets['PnB procesy']
    (max_row, max_col) = result.shape
    column_settings = [{'header': column} for column in result.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    writer.close()
