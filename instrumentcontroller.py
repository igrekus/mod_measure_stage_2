import ast
import time

import numpy as np

from collections import defaultdict
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal

from instr.instrumentfactory import mock_enabled, GeneratorFactory, SourceFactory, MultimeterFactory, AnalyzerFactory
from measureresult import MeasureResult
from util.file import load_ast_if_exists, pprint_to_file


GIGA = 1_000_000_000
MEGA = 1_000_000
KILO = 1_000
MILLI = 1 / 1_000


class InstrumentController(QObject):
    pointReady = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent=parent)

        addrs = load_ast_if_exists('instr.ini', default={
            'Анализатор': 'GPIB1::18::INSTR',
            'P LO': 'GPIB1::6::INSTR',
            'Источник': 'GPIB1::3::INSTR',
            'Мультиметр': 'GPIB1::22::INSTR',
        })

        self.requiredInstruments = {
            'Анализатор': AnalyzerFactory(addrs['Анализатор']),
            'P LO': GeneratorFactory(addrs['P LO']),
            'Источник': SourceFactory(addrs['Источник']),
            'Мультиметр': MultimeterFactory(addrs['Мультиметр']),
        }

        self.deviceParams = {
            'Модулятор': {
                'F': 1,
            },
        }

        self.secondaryParams = load_ast_if_exists('params.ini', default={
            'Plo': -5.0,
            'Flo_min': 0.05,
            'Flo_max': 6.05,
            'Flo_delta': 1.0,
            'is_Flo_div2': False,
            'Fmod': 1.0,   # MHz
            'Umod_min': 5,   # %
            'Umod_max': 100,   # %
            'Umod_delta': 5,   # %
            'Uoffs': 250,   # mV
            'Usrc': 5.0,
            'sa_rlev': 10.0,
            'sa_scale_y': 10.0,
            'sa_span': 10.0,   # MHz
            'sa_avg_state': False,
            'sa_avg_count': 16,
        })

        self._percent_to_db = load_ast_if_exists('modulate.ini', default={
            5: -20.77,
            10: -14.75,
            15: -11.23,
            20: -8.73,
            25: -6.79,
            30: -5.27,
            35: -3.93,
            40: -2.77,
            45: -1.74,
            50: -0.83,
            55: 0.0,
            60: 0.75,
            65: 1.45,
            70: 2.09,
            75: 2.69,
            80: 3.25,
            85: 3.78,
            90: 4.27,
            95: 4.74,
            100: 5.19,
        })

        self._calibrated_pows_lo = load_ast_if_exists('cal_lo.ini', default={})
        self._calibrated_pows_rf = load_ast_if_exists('cal_rf.ini', default={})

        self._instruments = dict()
        self.found = False
        self.present = False
        self.hasResult = False
        self.only_main_states = False

        self.result = MeasureResult()

    def __str__(self):
        return f'{self._instruments}'

    def connect(self, addrs):
        print(f'searching for {addrs}')
        for k, v in addrs.items():
            self.requiredInstruments[k].addr = v
        self.found = self._find()

    def _find(self):
        self._instruments = {
            k: v.find() for k, v in self.requiredInstruments.items()
        }
        return all(self._instruments.values())

    def check(self, token, params):
        print(f'call check with {token} {params}')
        device, secondary = params
        self.present = self._check(token, device, secondary)
        print('sample pass')

    def _check(self, token, device, secondary):
        print(f'launch check with {self.deviceParams[device]} {self.secondaryParams}')
        self._init()
        return True

    def calibrate(self, token, params):
        print(f'call calibrate with {token} {params}')
        return self._calibrate(token, self.secondaryParams)

    def _calibrateLO(self, token, secondary):
        print('run calibrate LO with', secondary)

        gen_lo = self._instruments['P LO']
        sa = self._instruments['Анализатор']

        secondary = self.secondaryParams

        lo_pow = secondary['Plo']
        lo_f_start = secondary['Flo_min'] * GIGA
        lo_f_end = secondary['Flo_max'] * GIGA
        lo_f_step = secondary['Flo_delta'] * GIGA

        lo_f_is_div2 = secondary['is_Flo_div2']

        sa_rlev = secondary['sa_rlev']
        sa_scale_y = secondary['sa_scale_y']
        sa_span = secondary['sa_span'] * MEGA

        freq_lo_values = [round(x, 3) for x in
                          np.arange(start=lo_f_start, stop=lo_f_end + 0.0001, step=lo_f_step)]

        sa.send(':CAL:AUTO OFF')
        sa.send(':SENS:FREQ:SPAN 1MHz')
        sa.send(f'DISP:WIND:TRAC:Y:RLEV 10')
        sa.send(f'DISP:WIND:TRAC:Y:PDIV 5')

        gen_lo.send(f':OUTP:MOD:STAT OFF')

        sa.send(':CALC:MARK1:MODE POS')
        sa.send(f':SENS:FREQ:SPAN {sa_span}Hz')
        sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
        sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')

        result = defaultdict(dict)
        gen_lo.send(f'SOUR:POW {lo_pow}dbm')

        for freq in freq_lo_values:

            freq_gen = freq
            if lo_f_is_div2:
                freq_gen *= 2

            if token.cancelled:
                gen_lo.send(f'OUTP:STAT OFF')
                time.sleep(0.5)

                gen_lo.send(f'SOUR:POW {lo_pow}dbm')

                gen_lo.send(f'SOUR:FREQ {lo_f_start}GHz')
                raise RuntimeError('calibration cancelled')

            gen_lo.send(f'SOUR:FREQ {freq_gen}Hz')

            gen_lo.send(f'OUTP:STAT ON')
            gen_lo.send(f':RAD:ARB ON')

            if not mock_enabled:
                time.sleep(0.5)

            sa.send(f':SENSe:FREQuency:CENTer {freq_gen}Hz')

            if not mock_enabled:
                time.sleep(0.5)

            sa.send(f':CALCulate:MARKer1:X {freq_gen}Hz')
            pow_read = float(sa.query(':CALCulate:MARKer:Y?'))
            loss = abs(lo_pow - pow_read)
            if mock_enabled:
                loss = 10

            print('loss: ', loss)
            result[lo_pow][freq_gen] = loss

        result = {k: v for k, v in result.items()}
        pprint_to_file('cal_lo.ini', result)

        gen_lo.send(f'OUTP:STAT OFF')
        sa.send(':CAL:AUTO ON')
        self._calibrated_pows_lo = result
        return True

    def _calibrateRF(self, token, secondary):
        print('run empty calibrate RF')

        result = dict()
        pprint_to_file('cal_rf.ini', result)

        self._calibrated_pows_rf = result
        return True

    def measure(self, token, params):
        print(f'call measure with {token} {params}')
        device, _ = params
        try:
            self.result.set_secondary_params(self.secondaryParams)
            self._measure(token, device)
            # self.hasResult = bool(self.result)
            self.hasResult = True  # HACK
        except RuntimeError as ex:
            print('runtime error:', ex)

    def _measure(self, token, device):
        param = self.deviceParams[device]
        secondary = self.secondaryParams
        print(f'launch measure with {token} {param} {secondary}')

        self._clear()
        self._measure_s_params(token, param, secondary)
        return True

    def _clear(self):
        self.result.clear()

    def _init(self):
        self._instruments['P LO'].send('*RST')
        self._instruments['Источник'].send('*RST')
        self._instruments['Мультиметр'].send('*RST')
        self._instruments['Анализатор'].send('*RST')

    def _measure_s_params(self, token, param, secondary):

        def set_read_marker(freq):
            sa.send(f':CALCulate:MARKer1:X {freq}Hz')
            if not mock_enabled:
                time.sleep(0.01)
            return float(sa.query(':CALCulate:MARKer:Y?'))

        gen_lo = self._instruments['P LO']
        src = self._instruments['Источник']
        mult = self._instruments['Мультиметр']
        sa = self._instruments['Анализатор']

        lo_pow = secondary['Plo']
        lo_f_start = secondary['Flo_min'] * GIGA
        lo_f_end = secondary['Flo_max'] * GIGA
        lo_f_step = secondary['Flo_delta'] * GIGA

        lo_f_is_div2 = secondary['is_Flo_div2']

        mod_f = secondary['Fmod'] * MEGA
        mod_u_offs = secondary['Uoffs'] * MILLI
        mod_f_offs_0 = 0.5 * MEGA  # MHz
        mod_u_min = secondary['Umod_min']
        mod_u_max = secondary['Umod_max']
        mod_u_delta = secondary['Umod_delta']

        src_u = secondary['Usrc']
        src_i_max = 200   # mA

        sa_rlev = secondary['sa_rlev']
        sa_scale_y = secondary['sa_scale_y']
        sa_span = secondary['sa_span'] * MEGA
        sa_avg_state = 'ON' if secondary['sa_avg_state'] else 'OFF'
        sa_avg_count = secondary['sa_avg_count']

        mod_u_values = [
            round(x, 3)for x in
            np.arange(start=mod_u_min, stop=mod_u_max + 0.0002, step=mod_u_delta)
        ]

        freq_lo_values = [
            round(x, 3) for x in
            np.arange(start=lo_f_start, stop=lo_f_end + 0.0001, step=lo_f_step)
        ]

        waveform_filename = 'WFM1:SINE_TEST_WFM'

        gen_lo.send(f':OUTP:MOD:STAT OFF')
        gen_lo.send(f':RAD:ARB OFF')
        gen_lo.send(f':RAD:ARB:WAV "{waveform_filename}"')
        gen_lo.send(f':RAD:ARB:BASE:FREQ:OFFS {mod_f + mod_f_offs_0}')
        gen_lo.send(f':DM:IQAD:EXT:COFF {mod_u_offs}')
        gen_lo.send(f':DM:IQAD:EXT:IQAT 0db')
        gen_lo.send(f':DM:IQAD ON')
        gen_lo.send(f':DM:STAT ON')

        gen_lo.send(f'SOUR:POW {lo_pow}dbm')

        src.send(f'APPLY p6v,{src_u}V,{src_i_max}mA')

        sa.send(':CAL:AUTO OFF')
        sa.send(f':SENS:FREQ:SPAN {sa_span}')
        sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
        sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')
        sa.send(f'AVER:COUNT {sa_avg_count}')
        sa.send(f'AVER {sa_avg_state}')
        sa.send(':CALC:MARK1:MODE POS')

        src.send('OUTPut ON')

        gen_lo.send(f'OUTP:STAT ON')
        gen_lo.send(f':RAD:ARB ON')

        if mock_enabled:
            with open('./mock_data/-5_16avg.txt', mode='rt', encoding='utf-8') as f:
                index = 0
                mocked_raw_data = ast.literal_eval(''.join(f.readlines()))

        res = []
        for freq_lo in freq_lo_values:

            freq_sa = freq_lo
            if lo_f_is_div2:
                freq_lo *= 2

            pow_loss = self._calibrated_pows_lo.get(lo_pow, dict()).get(freq_lo, 0) / 2

            gen_lo.send(f'SOUR:FREQ {freq_lo}')
            gen_lo.send(f'SOUR:POW {lo_pow + pow_loss}dbm')

            time.sleep(0.5)

            for mod_u in mod_u_values:

                if token.cancelled:
                    gen_lo.send(f'OUTP:STAT OFF')
                    gen_lo.send(f':DM:STAT OFF')
                    gen_lo.send(f':RAD:ARB OFF')
                    if not mock_enabled:
                        time.sleep(0.5)
                    src.send('OUTPut OFF')

                    gen_lo.send(f'SOUR:POW {lo_pow}dbm')
                    gen_lo.send(f'SOUR:FREQ {lo_f_start}Hz')

                    sa.send(':CAL:AUTO ON')
                    raise RuntimeError('measurement cancelled')

                if not mock_enabled:
                    time.sleep(0.5)

                if sa_avg_state:
                    sa.send(f'AVER OFF')

                gen_lo.send(f':RAD:ARB:RSC {mod_u}')

                if sa_avg_state:
                    sa.send(f'AVER ON')

                sa.send(f':SENSe:FREQuency:CENTer {freq_sa}')

                if not mock_enabled:
                    time.sleep(1.5)

                if lo_f_is_div2:
                    f_out = freq_sa + mod_f
                    sa_p_out = set_read_marker(f_out)

                    f_carr = freq_sa
                    sa_p_carr = set_read_marker(f_carr)

                    f_sb = freq_sa - mod_f
                    sa_p_sb = set_read_marker(f_sb)

                    f_3_harm = freq_sa - 3 * mod_f
                    sa_p_3_harm = set_read_marker(f_3_harm)
                else:
                    f_out = freq_sa - mod_f
                    sa_p_out = set_read_marker(f_out)

                    f_carr = freq_sa
                    sa_p_carr = set_read_marker(f_carr)

                    f_sb = freq_sa + mod_f
                    sa_p_sb = set_read_marker(f_sb)

                    f_3_harm = freq_sa + 3 * mod_f
                    sa_p_3_harm = set_read_marker(f_3_harm)

                # lo_p_read = float(gen_lo.query('SOUR:POW?'))
                # lo_f_read = float(gen_lo.query('SOUR:FREQ?'))

                src_u_read = src_u
                src_i_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

                raw_point = {
                    'lo_p': lo_pow,
                    'lo_f': freq_lo,
                    'mod_u': mod_u,   # in %
                    'mod_u_db': self._percent_to_db.get(mod_u, 0),   # in power values via % <-> db table
                    'src_u': src_u_read,   # power source voltage as set in GUI
                    'src_i': src_i_read,
                    'sa_p_out': sa_p_out,
                    'sa_p_carr': sa_p_carr,
                    'sa_p_sb': sa_p_sb,
                    'sa_p_3_harm': sa_p_3_harm,
                    'loss': pow_loss,
                }

                if mock_enabled:
                    raw_point = mocked_raw_data[index]
                    index += 1

                print(raw_point)
                self._add_measure_point(raw_point)
                res.append(raw_point)

        gen_lo.send(f'OUTP:STAT OFF')
        gen_lo.send(f':DM:STAT OFF')
        gen_lo.send(f':RAD:ARB OFF')
        if not mock_enabled:
            time.sleep(0.5)
        src.send('OUTPut OFF')

        gen_lo.send(f'SOUR:POW {lo_pow}dbm')
        gen_lo.send(f'SOUR:FREQ {lo_f_start}Hz')

        sa.send(':CAL:AUTO ON')

        if not mock_enabled:
            with open('out.txt', mode='wt', encoding='utf-8') as f:
                f.write(str(res))

        return res

    def _add_measure_point(self, data):
        print('measured point:', data)
        self.result.add_point(data)
        self.pointReady.emit()

    def saveConfigs(self):
        pprint_to_file('params.ini', self.secondaryParams)

    @pyqtSlot(dict)
    def on_secondary_changed(self, params):
        self.secondaryParams = params

    @property
    def status(self):
        return [i.status for i in self._instruments.values()]
