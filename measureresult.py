import os
import datetime

from collections import defaultdict
from subprocess import Popen
from textwrap import dedent

import pandas as pd

from util.file import load_ast_if_exists, pprint_to_file
from util.const import *


class MeasureResult:
    def __init__(self):
        self._secondaryParams = None
        self._raw = list()
        self._report = dict()
        self._processed = list()
        self._processed_cutoffs = list()
        self.ready = False

        self.data1 = defaultdict(list)
        self.data2 = defaultdict(list)

        self.adjustment = load_ast_if_exists('adjust.ini', default=None)

    def __bool__(self):
        return self.ready

    def _process(self):
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
            cutoffs[1].append([f_lo, cutoff_point])

        self.data2 = cutoffs
        self._processed_cutoffs = cutoffs
        self.ready = True

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
            point = self.adjustment[len(self._processed)]
            kp += point['kp']

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

        lo_f_label = lo_f
        self.data1[lo_f_label].append([p_in_db, sa_p_out])
        self._processed.append({**self._report})

    def clear(self):
        self._secondaryParams.clear()
        self._raw.clear()
        self._report.clear()
        self._processed.clear()
        self._processed_cutoffs.clear()

        self.data1.clear()
        self.data2.clear()

        self.ready = False

    def set_secondary_params(self, params):
        self._secondaryParams = dict(**params)

    def add_point(self, data):
        self._raw.append(data)
        self._process_point(data)

    def save_adjustment_template(self):
        if self.adjustment is None:
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
        # TODO implement
        device = 'mod'
        path = 'xlsx'
        if not os.path.isdir(f'{path}'):
            os.makedirs(f'{path}')
        file_name = f'./{path}/{device}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
        df = pd.DataFrame(self._processed)

        df.columns = [
            'Pгет, дБм', 'Fгет, ГГц', 'Pпот, дБ',
            'Pвх, %', 'Pвх, дБм',
            'Pвых, дБм', 'Pнес, дБм', 'Pбок, дБм', 'P3г, дБм',
            'Кп, дБ',
            'Uпит, В', 'Iпит, мА',
        ]
        df.to_excel(file_name, engine='openpyxl', index=False)

        full_path = os.path.abspath(file_name)
        Popen(f'explorer /select,"{full_path}"')
