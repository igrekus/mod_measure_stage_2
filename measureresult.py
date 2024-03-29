import os
import openpyxl
import random

from collections import defaultdict
from textwrap import dedent

import pandas as pd

from instr.const import *
from forgot_again.file import load_ast_if_exists, pprint_to_file, make_dirs, open_explorer_at
from forgot_again.string import now_timestamp


class MeasureResult:
    def __init__(self):
        self._primary_params = None
        self._secondaryParams = None
        self._raw = list()
        self._report = dict()
        self._processed = list()
        self._processed_cutoffs = list()
        self.ready = False

        self.data1 = defaultdict(list)
        self.data2 = defaultdict(list)

        self.adjustment = load_ast_if_exists('adjust.ini', default=None)
        self._table_header = list()
        self._table_data = list()

    def __bool__(self):
        return self.ready

    def process(self):

        cutoff_level = -1
        cutoffs = {1: []}

        for f_lo, datas in self.data1.items():
            reference = datas[0][1]
            cutoff_point = 0
            cutoff_idx = 0
            for idx, pair in enumerate(datas):
                pow_in, pow_out = pair
                if reference - pow_out > abs(cutoff_level):
                    cutoff_idx = idx
                    cutoff_point = pow_in
                    break
            else:
                cutoff_point = pow_in
            cutoffs[1].append([f_lo, cutoff_point])

        self.data2 = cutoffs
        self._processed_cutoffs = cutoffs

        self.ready = True
        self._prepare_table_data()

    def _process_point(self, data):

        lo_p = data['lo_p']
        lo_f = data['lo_f']

        src_u = data['src_u']
        src_i = data['src_i'] / MILLI

        loss = data['loss']
        p_in = data['mod_u']
        p_in_db = data['mod_u_db']

        sa_p_out = data['sa_p_out'] + loss
        sa_p_carr = data['sa_p_carr'] + loss
        sa_p_sb = data['sa_p_sb'] + loss
        sa_p_3_harm = data['sa_p_3_harm'] + loss

        kp = sa_p_out - p_in_db

        if self.adjustment is not None:
            try:
                point = self.adjustment[len(self._processed)]
                kp += point['kp']
            except LookupError:
                pass

        self._report = {
            'lo_p': lo_p,
            'lo_f': round(lo_f / GIGA, 3),
            'lo_p_loss': loss,
            'p_in': round(p_in, 2),
            'p_in_db': round(p_in_db, 2),

            'p_out': round(sa_p_out, 2),
            'p_carr': round(sa_p_carr, 2),
            'p_sb': round(sa_p_sb, 2),
            'p_3_harm': round(sa_p_3_harm, 2),

            'kp': round(kp, 2),

            'src_u': src_u,
            'src_i': round(src_i, 2),
        }

        lo_f_label = lo_f / GIGA
        self.data1[lo_f_label].append([p_in_db, kp])
        self._processed.append({**self._report})

    def clear(self):
        self._secondaryParams.clear()
        self._raw.clear()
        self._report.clear()
        self._processed.clear()
        self._processed_cutoffs.clear()

        self.data1.clear()
        self.data2.clear()

        self.adjustment = load_ast_if_exists(self._primary_params.get('adjust', ''), default={})

        self.ready = False

    def set_secondary_params(self, params):
        self._secondaryParams = dict(**params)

    def set_primary_params(self, params):
        self._primary_params = dict(**params)

    def add_point(self, data):
        self._raw.append(data)
        self._process_point(data)

    def save_adjustment_template(self):
        if not self.adjustment:
            print('measured, saving template')
            self.adjustment = [{
                'lo_p': p['lo_p'],
                'lo_f': p['lo_f'],
                'kp': 0,
            } for p in self._processed]
            pprint_to_file('adjust.ini', self.adjustment)

    @property
    def report(self):
        return dedent("""        Генератор:
        Pгет, дБм={lo_p}
        Fгет, ГГц={lo_f:0.2f}
        Pпот, дБ={lo_p_loss:0.2f}
        Pвх, %={p_in:0.2f}
        Pвх, дБ={p_in_db:0.2f}

        Источник питания:
        U, В={src_u}
        I, мА={src_i}

        Анализатор:
        Pвых, дБм={p_out:0.3f}
        Pнес, дБм={p_carr:0.3f}
        Pбок, дБм={p_sb}
        P3г, дБм={p_3_harm}

        Расчётные параметры:
        Кп, дБ={kp}
        """.format(**self._report))

    def export_excel(self):
        device = 'mod'
        path = 'xlsx'

        make_dirs(f'{path}')
        file_name = f'./{path}/{device}-kp-{now_timestamp()}.xlsx'

        df = pd.DataFrame(self._processed)
        df.columns = [
            'Pгет, дБм', 'Fгет, ГГц', 'Pпот, дБ',
            'Pвх, %', 'Pвх, дБм',
            'Pвых, дБм', 'Pнес, дБм', 'Pбок, дБм', 'P3г, дБм',
            'Кп, дБ',
            'Uпит, В', 'Iпит, мА',
        ]
        df.to_excel(file_name, engine='openpyxl', index=False)

        self._export_cutoff()

        open_explorer_at(os.path.abspath(file_name))

    def _export_cutoff(self):
        device = 'mod'
        path = 'xlsx'
        file_name = f'./{path}/{device}-cutoff-{now_timestamp()}.xlsx'
        df = pd.DataFrame(self._processed_cutoffs[1], columns=['lo_f', 'cutoff'])

        df.columns = ['Fгет, ГГц', 'P1дБвх, дБм']

        df.to_excel(file_name, engine='openpyxl', index=False)

    def _prepare_table_data(self):
        table_file = self._primary_params.get('result', '')

        if not os.path.isfile(table_file):
            return

        wb = openpyxl.load_workbook(table_file)
        ws = wb.active

        rows = list(ws.rows)
        self._table_header = [row.value for row in rows[0][1:]]

        gens = [
            [rows[1][j].value, rows[2][j].value, rows[3][j].value]
            for j in range(1, ws.max_column)
        ]

        self._table_data = [self._gen_value(col) for col in gens]

    def _gen_value(self, data):
        if not data:
            return '-'
        if '-' in data:
            return '-'
        span, step, mean = data
        start = mean - span
        stop = mean + span
        if span == 0 or step == 0:
            return mean
        return round(random.randint(0, int((stop - start) / step)) * step + start, 2)

    def get_result_table_data(self):
        return list(self._table_header), list(self._table_data)
