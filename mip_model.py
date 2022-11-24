from typing import List

import gurobipy as gp
from gurobipy import GRB
from openpyxl.styles import Font

# import odfpy
# import openpyxl

import itertools
import pandas as pd
from collections import namedtuple
from openpyxl.styles import Font

hora_elementar = [(8, 9), (9, 10), (10, 11), (11, 12), (14, 16), (16, 17), (17, 18)]
horarios = [(8, 10), (10, 12), (14, 16), (16, 18)]  # 2h
horarios += [(14, 17)]  # 3
horarios += [(14, 18)]  # 4
horarios += [(8, 9), (9, 10), (10, 11), (11, 12), (14, 15)]  # 1
horarios = sorted(horarios)

Horario = namedtuple("Horario", "inicio fim")
horarios = [Horario(*x) for x in horarios]

dias = ['seg', 'ter', 'quar', 'quin', 'sex']
dias_dic = {dias[i]: i for i in range(len(dias))}
dias_patter = {1: [set([d]) for d in dias]}
dias_patter[2] = [set(d) for d in itertools.combinations(dias, 2)]

slots_pool = list(itertools.product(horarios, dias_patter[1]))
slots_pool += list(itertools.product(horarios, dias_patter[2]))
# slots_pool += list(itertools.product(horarios, dias_patter[3]))
Slot = namedtuple("Slot", "hora dias")
slots_pool = [Slot(*x) for x in slots_pool]


def slots(disc_ch, proibir_dias, proibir_horarios, fixar_dias, fixar_hora, prof_externo=False):
    """
    Lista os slots no pool compatíveis com as restrições
    :param disc_ch:
    :param proibir_dias:
    :param proibir_horarios:
    :param fixar_dias:
    :param fixar_hora:
    :param prof_externo:
    :return: lista de slots no pool compatíveis com as restrições
    """
    l: list[int] = []
    for i, s in enumerate(slots_pool):
        time = (s.hora.fim - s.hora.inicio) * len(s.dias) * 15
        if disc_ch <= time <= disc_ch + 15 \
                and s.dias.isdisjoint(proibir_dias) \
                and not horario_colide(s.hora, proibir_horarios):
            # if s.hora.fim - s.hora.inicio >= 4 and not prof_externo:
            #     continue
            if s.hora.fim - s.hora.inicio >= 4 and fixar_hora[0] == 0:
                # quatro horas, só com se o horario for prefixado
                continue
            if len(fixar_dias) == 0 or s.dias.issubset(fixar_dias):
                if fixar_hora[0] == 0 or (fixar_hora[0] <= s.hora.inicio and s.hora.fim <= fixar_hora[1]):
                    l.append(i)
    return l


def read_input(filepath="input.ods"):
    disc_df = pd.read_excel(filepath, sheet_name="disciplinas")
    disc_df['cod'] = disc_df['cod'].str.strip()
    if disc_df['cod'].duplicated().any():
        print('Códigos de disciplinas duplicados:')
        print(disc_df[disc_df['cod'].duplicated()])
        exit(1)
    disc_df.set_index('cod', inplace=True)
    disc_df['grades'] = disc_df['grades'].apply(lambda a: a.split())
    disc_df['semestres'] = disc_df['semestres'].apply(
        lambda a: [int(x) for x in a.split()] if isinstance(a, str) else [a])
    disc_df['semestres'] = disc_df.apply(
        lambda a: a['semestres'] if len(a['semestres']) == len(a['grades']) else a['semestres'] + [
            a['semestres'][-1]] * (len(a['grades']) - len(a['semestres'])), axis=1)

    disc_df['n_turmas'] = disc_df['n_turmas'].fillna(1).astype(int)
    disc_df['cod_lab'].fillna('', inplace=True)
    disc_df['proibir_dias'].fillna('', inplace=True)
    disc_df['proibir_dias'] = disc_df['proibir_dias'].apply(lambda a: set(a.split()))
    disc_df['proibir_horario'].fillna('0-0', inplace=True)
    disc_df['proibir_horario'] = disc_df['proibir_horario'].apply(
        lambda a: (int(a.split('-')[0]), int(a.split('-')[1])))
    disc_df['fixar_dias'].fillna('', inplace=True)
    disc_df['fixar_dias'] = disc_df['fixar_dias'].apply(lambda a: set(a.split()))
    disc_df['fixar_prof'].fillna('', inplace=True)
    disc_df['fixar_horario'].fillna('0-0', inplace=True)
    disc_df['fixar_horario'] = disc_df['fixar_horario'].apply(
        lambda a: (int(a.split('-')[0]), int(a.split('-')[1])))

    prof_df = pd.read_excel(filepath, sheet_name="professores", index_col='cod')
    prof_df['disciplinas'] = prof_df['disciplinas'].apply(lambda a: set(a.split()))
    prof_df['dias'] = prof_df[dias].apply(lambda r: set([d for d in dias if r[d] == 1]), axis=1)
    prof_df['ch_max'] = prof_df['ch_max'].fillna(195).astype(int)
    prof_df['ch_min'] = prof_df['ch_min'].fillna(0).astype(int)
    prof_df['dias_max'] = prof_df['dias_max'].fillna(5).astype(int)
    prof_df['priorizar'] = prof_df['priorizar'].fillna('h')
    prof_compativeis = {x: set() for x in disc_df.index}
    turmas_nao_cadastradas = set()

    for k, prof_row in prof_df.iterrows():
        for turma_cod in prof_row['disciplinas']:
            if turma_cod in prof_compativeis:
                prof_compativeis[turma_cod].add(k)
            else:
                turmas_nao_cadastradas.add(turma_cod)
    for turma_cod, row in disc_df.iterrows():
        if row['fixar_prof'] != '':
            prof_compativeis[turma_cod] = set([int(row['fixar_prof'])])

    dup = disc_df[disc_df['n_turmas'] > 1]
    disc_df.drop(dup.index, inplace=True)
    for k, r in dup.iterrows():
        for i in range(r['n_turmas']):
            tmp = r.copy()
            tmp['semestres'] = [tmp['semestres'][i]] if len(tmp['semestres']) > i else [
                tmp['semestres'][i % len(tmp['semestres'])]]
            tmp['grades'] = [tmp['grades'][i]] if len(tmp['grades']) > i else [tmp['grades'][i % len(tmp['grades'])]]
            tmp['turmas'] = 1
            if i == 0:
                disc_df.loc[k] = tmp
            else:
                turma_cod = k + f'_t{i + 1}'
                disc_df.loc[turma_cod] = tmp
                prof_compativeis[turma_cod] = prof_compativeis[k]

    if len(turmas_nao_cadastradas) > 0:
        print(
            f"\n[ATENÇÃO]: As seguintes {len(turmas_nao_cadastradas)}  disciplinas estão cadastradas no professores, mas não estão na tabela de turmas: \n",
            '\n'.join(str(x) for x in turmas_nao_cadastradas))

    ignore_disc = [[x, disc_df.loc[x]['nome']] for x in disc_df.index.values if len(prof_compativeis[x]) == 0]
    if len(ignore_disc) > 0:
        print(
            f'\n[ATENÇÃO]: As seguintes {len(ignore_disc)}  turmas não possuem professores compatíveis e serão ignoradas: \n',
            '\n'.join(str(x) for x in ignore_disc))
    ignore_prof = [[x, prof_df.loc[x]['nome']] for x in prof_df.index if
                   len(set(prof_df.loc[x]['disciplinas']).intersection(set(disc_df.index))) == 0]
    if len(ignore_prof) > 0:
        print(
            f'\n[ATENÇÃO]: As seguintes {len(ignore_prof)}  professores não possuem turmas compatíveis e serão ignorados: \n{ignore_prof}')
    if len(ignore_prof) > 0:
        cods, null = zip(*ignore_prof)
        prof_df.drop(list(cods), inplace=True)
    if len(ignore_disc) > 0:
        cods, null = zip(*ignore_disc)
        disc_df.drop(list(cods), inplace=True)

    grades_unique = set()
    for g in disc_df['grades']:
        for i in g:
            grades_unique.add(i)

    semestres_unique = set()
    for sem in disc_df['semestres']:
        for i in sem:
            semestres_unique.add(i)

    print("\nDistribuição por Grade por Semestre")
    for g in grades_unique:
        print("GRADE: ", g)
        for sem in semestres_unique:
            exp = [d for d in disc_df.index if (sem, g) in zip(disc_df.loc[d]['semestres'], disc_df.loc[d]['grades'])]
            ch = disc_df.loc[exp]['ch'].sum()
            print(f'Semestre: {sem}  CH:{ch}  Qtd:{len(exp)}')

    return disc_df, prof_df, prof_compativeis, ignore_disc, ignore_prof, grades_unique, semestres_unique


def make_prof_output(prof_timetable, prof_df, disc_df, ch):
    data = []
    for prof_cod in prof_timetable.keys():
        prof_nome = prof_df.loc[prof_cod]["nome"].strip()
        for h in horarios:
            row = [prof_nome]
            if (ch[prof_cod].X > prof_df.loc[prof_cod]['ch_max'] + 1e-3):
                row.append(f'{round(ch[prof_cod].X)} (> {prof_df.loc[prof_cod]["ch_max"]})')
            elif (ch[prof_cod].X < prof_df.loc[prof_cod]['ch_min'] - 1e-3):
                row.append(f'{round(ch[prof_cod].X)} (< {prof_df.loc[prof_cod]["ch_min"]})')
            else:
                row.append(round(ch[prof_cod].X))
            row.append(f'{h[0]}-{h[1]}')
            empty = True
            for d in dias:
                disc_cod = prof_timetable[prof_cod][d, h]
                if disc_cod == '':
                    row.append('')
                else:
                    empty = False
                    txt = f'{disc_df.loc[disc_cod]["nome"].strip()}{disc_df.loc[disc_cod]["grades"]}'
                    if disc_cod.count('_t') == 1:
                        txt += f"(turma {disc_cod.split('_t')[1]})"
                    if d not in prof_df.loc[prof_cod]['dias']:
                        txt += "(*)"
                    if h[1] - h[0] == 1:
                        txt += "(**)"
                    row.append(txt)
            if not empty:
                data.append(row)
    por_professor_df = pd.DataFrame(data=data, columns=['professor', 'CH', 'horário'] + dias)
    por_professor_df.set_index(['professor', 'CH', 'horário'], inplace=True)
    return por_professor_df


def make_semestre_output(sem_timetable, prof_df, disc_df, grade):
    data = []
    for g, s in sem_timetable.keys():
        if g != grade:
            continue
        for h in horarios:
            row = [s]
            row.append(f'{h[0]}-{h[1]}')
            empty = True
            for d in dias:
                sol = sem_timetable[g, s][d, h]
                if sol == '':
                    row.append('')
                elif not isinstance(sol, list):
                    empty = False
                    prof_nome = prof_df.loc[sol[1]]["nome"].split("–")[0].strip()
                    txt = f'{disc_df.loc[sol[0]]["nome"].strip()}\n{prof_nome}'
                    disc_cod = sol[0]
                    if disc_cod.count('_t') == 1:
                        txt += f"(turma {disc_cod.split('_t')[1]})"
                    if h[1] - h[0] == 1:
                        txt += "(**)"
                    row.append(txt)
                else:
                    text = 'Colisão'
                    empty = False
                    for x in sol:
                        prof_nome = prof_df.loc[x[1]]["nome"].split("–")[0].strip()
                        text += f'\n{disc_df.loc[x[0]]["nome"].strip()}\n{prof_nome}'
                    row.append(text)

            # print(row)
            if not empty:
                data.append(row)
    por_semestre_df = pd.DataFrame(data=data, columns=['semestre', 'horário'] + dias)
    por_semestre_df.set_index(['semestre', 'horário'], inplace=True)
    return por_semestre_df


def slot_penalty(slot_cod):
    slot = slots_pool[slot_cod]
    p = 0
    # if slot.hora.fim >= 18:
    #     p = 1e6
    if slot.hora.fim >= 17:
        p = 1e2
    elif slot.hora.fim >= 12:
        p = 2
    elif slot.hora.fim == 9:
        p = 1

    if slot.hora.fim >= 16 and 'sex' in slot.dias:
        p += 1e1

    if len(slot[1]) == 2:
        #dias consecutivos
        l = list(slot[1])
        p += abs(abs(dias_dic[l[1]] - dias_dic[l[0]]) - 2) * 1e1
    return p


# def slot_penalty(slot_cod, prof_dias, priorizar,x=''):
#     slot = slots_pool[slot_cod]
#
#     if x == 'd':
#         return len(slot.dias - prof_dias)
#
#     if priorizar == 'h':
#         p_h = 1e6
#         p_d = 1e3
#     else:
#         p_h = 1e3
#         p_d = 1e6
#     p = p_d * len(slot[1] - prof_dias)
#     # p=0
#     if slot[0][1] >= 18:
#         p += p_h
#     elif slot[0][1] >= 17:
#         p += 1e1
#     elif slot[0][1] >= 12:
#         p += 1e0
#     elif slot[0][1] == 9:
#         p += 1e1
#
#     if slot[0][1] >= 16 and 'sex' in slot[1]:
#         p += 1e2
#
#     if len(slot[1]) == 2:
#         l = list(slot[1])
#         p += abs(abs(dias_dic[l[1]] - dias_dic[l[0]]) - 2) * 1e1
#     return p


def horario_colide(h1, h2):
    if h2[0] <= h1[0] < h2[1]:
        return True
    if h1[0] <= h2[0] < h1[1]:
        return True
    return False


def model():
    disc_df, prof_df, prof_compativeis, ignore_disc, ignore_prof, grades_unique, semestres_unique = read_input(
        'input2022.ods')

    model = gp.Model()
    x = {}
    prof_dias_dic = prof_df['dias'].to_dict()
    disc_proibir_dia = disc_df['proibir_dias'].to_dict()
    disc_proibir_hora = disc_df['proibir_horario'].to_dict()
    disc_fixar_dia = disc_df['fixar_dias'].to_dict()
    disc_fixar_hora = disc_df['fixar_horario'].to_dict()
    for d in disc_df.index:
        for p in prof_compativeis[d]:
            prof_ext = prof_df.loc[p]['nome'].count('outro instituto') > 0
            priorizar = prof_df.loc[p]['priorizar']
            for s in slots(disc_df.loc[d]['ch'], disc_proibir_dia[d], disc_proibir_hora[d], disc_fixar_dia[d],
                           disc_fixar_hora[d], prof_ext):
                x[d, p, s] = model.addVar(obj=0, vtype='b', name=f'x_{d}_{p}_{s}')

    ch_max_pen = {p: model.addVar(obj=0, vtype=GRB.INTEGER, name=f'ch_{p}') for p in prof_df.index}
    ch_min_pen = {p: model.addVar(obj=0, vtype=GRB.INTEGER, name=f'ch_{p}') for p in prof_df.index if
                  prof_df.loc[p]['ch_min'] > 0}
    ch = {p: model.addVar(vtype=GRB.INTEGER, name=f'ch_{p}') for p in prof_df.index}
    ch_a = {p: model.addVar(obj=1, vtype=GRB.INTEGER, name=f'ch_a_{p}') for p in prof_df.index}
    ch_b = {p: model.addVar(obj=1, vtype=GRB.INTEGER, name=f'ch_b_{p}') for p in prof_df.index}

    colisao = {(h, d, s, g): model.addVar(obj=0, vtype=GRB.INTEGER, name='col') for h in hora_elementar for d in dias
               for s in semestres_unique for g in grades_unique}

    prof_dia = {(p, d): model.addVar(obj=0, vtype=GRB.BINARY, name=f'pd_{p}{d}') for p in prof_df.index for d in dias}

    model.update()
    # prof dias
    for p in prof_df.index:
        for dia in dias:
            exp = [x[d, p, s] for d, pp, s in x.keys() if pp == p and dia in slots_pool[s][1]]
            model.addConstr(sum(exp) <= 8 * prof_dia[p, dia])
    for p in prof_df.index:
        exp = [prof_dia[p, dia] for dia in dias]
        model.addConstr(sum(exp) <= prof_df.loc[p]['dias_max'])
        # model.addConstr(sum(exp) <= 2)

    media = sum(disc_df['ch']) / len(prof_df)
    print("Carga horária média Geral:", media)

    soma = sum(disc_df['ch'])
    prof_abaixo_media = prof_df[prof_df.ch_max < media]['ch_max']
    media = (soma - sum(prof_abaixo_media)) / (len(prof_df) - len(prof_abaixo_media))
    print("Carga horária média para quem não tem redução abaixo da média:", media)
    media = round(media)  # acelera o branch and bound
    # carga horaria
    for p in prof_df.index:
        pch_max = prof_df.loc[p]['ch_max']
        M = media if pch_max >= media else pch_max
        model.addConstr(ch[p] >= 1)
        model.addConstr(ch[p] - M <= ch_a[p])
        model.addConstr(M - ch[p] <= ch_b[p])
        model.addConstr(sum(x[d, p, s] * disc_df.loc[d]['ch'] for d, pp, s in x.keys() if p == pp) == ch[p], f'ch_{p}')
        model.addConstr(ch[p] <= pch_max + 15 * ch_max_pen[p])
        # model.addConstr(ch[p] <= 195)
        if prof_df.loc[p]['ch_min'] > 0:
            model.addConstr(ch[p] + 15 * ch_min_pen[p] >= prof_df.loc[p]['ch_min'])

    # cobertura das disciplinas
    for d in disc_df.index:
        exp = [x[d, p, s] for dd, p, s in x.keys() if dd == d]
        if len(exp) != 0:
            model.addConstr(sum(exp) == 1, f'disc_{d}')
    model.update()

    disc_sem_dic = disc_df['semestres'].to_dict()
    disc_grades_dic = disc_df['grades'].to_dict()
    disc_lab_dic = disc_df['cod_lab'].to_dict()
    lab_unique = disc_df[disc_df['cod_lab'] != '']['cod_lab'].unique()
    # timetable
    for h in hora_elementar:
        for dia in dias:
            # timetable prof
            for p in prof_df.index:
                exp = [x[d, p, s] for d, pp, s in x.keys() if
                       pp == p and horario_colide(slots_pool[s][0], h) and dia in slots_pool[s][1]]
                if len(exp) > 0:
                    model.addConstr(sum(exp) <= 1, f'ttp_{p}_{h}{dia}')
            # timetable semestre
            for sem in semestres_unique:
                for g in grades_unique:
                    exp = [x[d, p, s] for d, p, s in x.keys() if
                           (sem, g) in zip(disc_sem_dic[d], disc_grades_dic[d]) and dia in slots_pool[s][1]
                           and horario_colide(slots_pool[s][0], h)]
                    if len(exp) > 0:
                        # model.addConstr(sum(exp) <= 1 , f'tts_{sem}_{h}{dia}')
                        model.addConstr(sum(exp) <= 1 + colisao[h, dia, sem, g], f'tts_{sem}{g}_{h}{dia}')
            # lab
            for lab in lab_unique:
                exp = [x[d, p, s] for d, p, s in x.keys() if
                       lab == disc_lab_dic[d] and dia in slots_pool[s][1] and horario_colide(slots_pool[s][0], h)]
                if len(exp) > 0:
                    model.addConstr(sum(exp) <= 1, f'ttl_{sem}{g}_{h}{dia}')

    # model.write('teste.lp')
    # model.setObjective(sum(ch[p]**2 for p in ch.keys()),GRB.MINIMIZE)
    model.update()
    # model.setObjective(model.getObjective() + sum([ch[p] ** 2 for p in prof_df.index]), GRB.MINIMIZE)

    model.setObjectiveN(sum(colisao.values()), 0, 7, name="colisão")
    model.setObjectiveN(sum(ch_max_pen.values()) + sum(ch_min_pen.values()), 1, 6, abstol=2, name="limite ch")
    model.setObjectiveN(sum(ch_a.values()) + sum(ch_b.values()), 2, 5, abstol=150, name="balanço ch")
    model.setObjectiveN(sum([prof_dia[p, d] for p, d in prof_dia.keys() if d not in prof_dias_dic[p]]), 3, 4,
                        name="dias convenience")
    # print([(p,d) for p, d in prof_dia.keys() if d not in prof_dias_dic[p]])
    model.setObjectiveN(
        sum([x[d, p, s] for d, p, s in x.keys() if disc_fixar_hora[d] == (0, 0) and slots_pool[s].hora.fim >= 18]), 4,
        3, name="hora 18 convenience")


    model.setObjectiveN(sum([x[d, p, s] * slot_penalty(s) for d, p, s in x.keys() if disc_fixar_hora[d] == (0, 0)]), 5,
                        2, reltol=.01, name="hora convenience")

    model.setObjectiveN(sum([prof_dia[p, d] for p, d in prof_dia.keys()]), 6, 1, name="prof dias")

    model.setParam('TimeLimit', 60 * 20)
    model.setParam('MIPGap', 0.015)
    # model.setParam('MIPFocus', 1)
    model.optimize()
    # if colisao.X > 0:
    #     print("Atenção: Distribuição inviável")

    make_output(ch, disc_df, disc_grades_dic, disc_sem_dic, grades_unique, ignore_disc, ignore_prof, prof_df,
                semestres_unique, x, prof_compativeis)


def make_output(ch, disc_df, disc_grades_dic, disc_sem_dic, grades_unique, ignore_disc, ignore_prof, prof_df,
                semestres_uniq, x, prof_compativeis):
    prof_timetable = {p: {(dia, h): '' for dia in dias for h in horarios} for p in prof_df.index}
    sem_timetable = {(g, s): {(dia, h): '' for dia in dias for h in horarios} for s in semestres_uniq for
                     g in grades_unique}
    for d, p, s in x.keys():
        if x[d, p, s].X == 1:
            if x[d, p, s].obj > 0:
                print('penalidade: ', d, p, s, x[d, p, s].obj)

            slot = slots_pool[s]
            # print(d, s, p, slot, sem)
            for dia in slots_pool[s][1]:
                prof_timetable[p][dia, slot[0]] = d
                for g, sem in zip(disc_grades_dic[d], disc_sem_dic[d]):
                    if sem_timetable[g, sem][dia, slot[0]] != '':
                        print('Colisão ', sem_timetable[g, sem][dia, slot[0]], (d, p))
                        if isinstance(sem_timetable[g, sem][dia, slot[0]], list):
                            sem_timetable[g, sem][dia, slot[0]].append((d, p))
                        else:
                            sem_timetable[g, sem][dia, slot[0]] = [sem_timetable[g, sem][dia, slot[0]], (d, p)]
                    else:
                        sem_timetable[g, sem][dia, slot[0]] = (d, p)
    por_professor_df = make_prof_output(prof_timetable, prof_df, disc_df, ch)
    por_semestre_df = {}
    for g in grades_unique:
        por_semestre_df[g] = make_semestre_output(sem_timetable, prof_df, disc_df, g)

    prof_pontos_df = make_pontos_prof(prof_df, x)
    disc_pontos_df = make_pontos_disc(disc_df, prof_df, x, prof_compativeis)
    dias_df = make_dias(prof_df, x)

    with pd.ExcelWriter('output.xlsx') as writer:

        por_professor_df.to_excel(writer, sheet_name='por professor')
        format_cells(writer.sheets['por professor'])

        dias_df.to_excel(writer, sheet_name='Dias')
        format_cells(writer.sheets['Dias'])

        for g in grades_unique:
            por_semestre_df[g].to_excel(writer, sheet_name=f'por semestre {g}')
            format_cells(writer.sheets[f'por semestre {g}'])
        if len(ignore_disc) > 0:
            ignore_df = pd.DataFrame(ignore_disc, columns=['cod', 'Nome'])
            ignore_df.set_index('cod')
            ignore_df.to_excel(writer, sheet_name='Diciplinas ignoradas')
            format_cells(writer.sheets['Diciplinas ignoradas'])
        if len(ignore_prof) > 0:
            ignore_df = pd.DataFrame(ignore_prof, columns=['cod', 'Nome'])
            ignore_df.set_index('cod')
            ignore_df.to_excel(writer, sheet_name='Professores ignorados')
            format_cells(writer.sheets['Professores ignorados'])

        prof_pontos_df.to_excel(writer, sheet_name='Pontos professores', index=None)
        format_cells(writer.sheets['Pontos professores'])
        disc_pontos_df.to_excel(writer, sheet_name='Pontos disciplinas', index=None)
        format_cells(writer.sheets['Pontos disciplinas'])


def make_pontos_prof(prof_df, x):
    data = []
    for p in prof_df.index:
        cont = 0
        for d, pp, s in x.keys():
            if pp == p:
                cont += 1 / (1 + x[d, p, s].obj)
        data.append([prof_df.loc[p]['nome'], round(cont, 2)])
    prof_pontos_df = pd.DataFrame(data=data, columns=['Professor', 'Pontos'])
    prof_pontos_df.sort_values('Pontos', inplace=True)
    return prof_pontos_df


def make_dias(prof_df, x):
    data = []
    for dia in dias:
        profs = set()
        for p in prof_df.index:
            for d, pp, s in x.keys():
                if pp == p and x[d, pp, s].X == 1 and dia in slots_pool[s][1]:
                    profs.add(prof_df.loc[p]['nome'])
        data.append([dia, "\n".join(sorted(list(profs)))])
    dias_df = pd.DataFrame(data=data, columns=['Dia', 'Professores'])
    dias_df.set_index(['Dia'], inplace=True)
    return dias_df


def make_pontos_disc(disc_df, prof_df, x, prof_compativeis):
    data = []
    for d in disc_df.index:
        cont = 0
        for dd, p, s in x.keys():
            if dd == d:
                cont += 1 / (1 + x[d, p, s].obj)
        data.append([d, disc_df.loc[d]['nome'], round(cont, 2),
                     "\n".join([prof_df.loc[p]['nome'] for p in prof_compativeis[d]])])
        prof_pontos_df = pd.DataFrame(data=data, columns=['cod', 'Nome', 'Pontos', 'Professores'])
        prof_pontos_df.sort_values('Pontos', inplace=True)
    return prof_pontos_df


def format_cells(ws):
    for column_cells in ws.columns:
        length = max(len(txt) for cell in column_cells for txt in str(cell.value).split('\n'))
        ws.column_dimensions[column_cells[0].column_letter].width = 10 + length * .8
    for row_cells in ws.rows:
        length = max(str(cell.value).count('\n') for cell in row_cells) + 1
        ws.row_dimensions[row_cells[0].row].height = length * 12
    for row_cells in ws.rows:
        for cell in row_cells:
            if str(cell.value).count("(**)") > 0:
                cell.style = "Note"
            if str(cell.value).count("Colisão") > 0:
                cell.style = "Bad"
            if str(cell.value).count("(*)") > 0:
                cell.style = "Bad"
            if str(cell.value).count("-18") > 0:
                cell.style = "Neutral"
            if str(cell.value).count(">") > 0 or str(cell.value).count("<") > 0:
                cell.style = "Neutral"


model()
