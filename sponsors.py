# -*- coding: utf-8 -*-

import os
import os.path as path
from itertools import chain

from errbot import BotPlugin, arg_botcmd
from tabulate import tabulate
from eptools.sponsors import (get_sponsor,
                              get_sponsors_ws_data,
                              create_sponsor_agreement,
                              )

# CNFG_DIR = path.dirname(__file__)
CNFG_DIR = path.join(os.environ.get('ERRBOT_CFG_DIR'),  'plugins', 'err-sponsors')
DATA_DIR = path.join(os.environ.get('ERRBOT_BASE_DIR'), 'data',    'err-sponsors')

CONFIG_TEMPLATE = {'GOOGLE_API_KEYFILE': path.join(CNFG_DIR, 'google_api_key.json'),
                   'SPONSORS_SHEET_KEY': '16ohl6y4n9RXfG5jizBYl1ns12UKFS3Crauyc1ZsP1G0', ## EP2017
                   #'SPONSORS_SHEET_KEY': '1Dbxy1a0c-IbXdxVXmbTeU6zE6AKV4N0xKK2GANSw9Mw',
                   'SPONSORS_SHEET_TAB': 'Form responses 1',
                   'INFO_COLUMNS': ('company', 'representative', 'email'),
                   'CONTRACTS_DIR': path.join(DATA_DIR, 'agreements'),
                   'TEMPLATE_FILE': {'eps': path.join(CNFG_DIR, 'sponsor_agreement_template_eps.tex'),
                                     'aps': path.join(CNFG_DIR, 'sponsor_agreement_template_aps.tex'),
                                     }
                   }


class SponsorsPlugin(BotPlugin):
    """ An Errbot plugin to generate sponsorship agreements for EuroPython."""

    def configure(self, configuration):
        if configuration is not None and configuration != {}:
            config = dict(chain(CONFIG_TEMPLATE.items(),
                                configuration.items()))
        else:
            config = CONFIG_TEMPLATE
        super(SponsorsPlugin, self).configure(config)
        self.startup()

    def _get_sender(self, message):
        """ Return a room ID to identify the feed reports destinations."""
        return message.frm if message.is_direct else message.to

    def get_configuration_template(self):
        return CONFIG_TEMPLATE

    def startup(self):
        contract_dir = self.config['CONTRACTS_DIR']
        if not path.exists(contract_dir):
            os.makedirs(contract_dir)

    def _sponsor_data(self, company):
        api_key_file = self.config['GOOGLE_API_KEYFILE']
        gdrive_sheet = self.config['SPONSORS_SHEET_KEY']

        sponsors = get_sponsors_ws_data(api_key_file=api_key_file,
                                        doc_key=gdrive_sheet)

        sponsor_data = self.pick_one_sponsor(sponsors, company)

        return sponsor_data

    @staticmethod
    def pick_one_sponsor(sponsors, sponsor_name):
        try:
            sponsor_data = get_sponsor(sponsor_name=sponsor_name,
                                       sponsors=sponsors,
                                       col_name='company')
        except:
            raise KeyError('Could not find data for sponsor {}.'.format(sponsor_name))
        else:
            if len(sponsor_data) != 1:
                raise KeyError("Found more than one sponsor: {}.".format(sponsor_data.to_json()))

            return sponsor_data

    @arg_botcmd('company', type=str)
    def sponsor_info(self, msg, company):
        """Give details about the sponsor."""
        info_cols = self.config['INFO_COLUMNS']

        sponsor_data = self._sponsor_data(company)

        return tabulate([(col, sponsor_data[col].values[0])
                         for col in info_cols],
                        tablefmt='pipe')

    def _get_contract_template(self, contract_type):
        templates = self.config['TEMPLATE_FILE']
        if isinstance(templates, str):
            return templates
        elif isinstance(templates, dict):
            return templates[contract_type]
        else:
            raise KeyError('Could not find a contract for contract_type {}.'.format(contract_type))

    @arg_botcmd('-c', dest='company', type=str)
    @arg_botcmd('-t', dest='contract_type', type=str, default='eps')
    def sponsor_agreement(self, msg, company, contract_type=None):
        output_dir = self.config['CONTRACTS_DIR']
        contract_template = self._get_contract_template(contract_type)

        try:
            sponsor_data = self._sponsor_data(company)
        except Exception as exc:
            self.log.exception(str(exc))
            return str(exc)
        else:
            fpath = create_sponsor_agreement(sponsor_data,
                                             template_file=contract_template,
                                             field_name='company',
                                             output_dir=output_dir)

            stream = self.send_stream_request(self._get_sender(msg),
                                              open(fpath, 'rb'),
                                              name=path.basename(fpath),
                                              size=path.getsize(fpath),
                                              stream_type='document')
