import json

from sqlalchemy import Column, Text
from sqlalchemy.ext.declarative import declared_attr
from sqlalchemy.orm import synonym


class MetaMixin(object):
    _metas = Column('metas', Text, nullable=True)

    def _get_metas(self):
        if self._metas is not None:
            return json.loads(self._metas)

    def _set_metas(self, metas):
        self._metas = json.dumps(metas)

    @declared_attr
    def metas(cls):
        return synonym("_metas", descriptor=property(cls._get_metas, cls._set_metas))
