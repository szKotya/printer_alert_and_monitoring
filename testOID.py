from pysnmp.hlapi.v3arch.asyncio import (
    get_cmd, SnmpEngine, CommunityData, UdpTransportTarget,
    ContextData, ObjectType, ObjectIdentity
)
import asyncio
import codecs
import sys
import locale
async def GetTonersStatus():
    ip = '192.168.0.248'
    community = 'public'
    test_oid = '1.3.6.1.2.1.43.16.5.1.2.1.1'

    target = await UdpTransportTarget.create((ip, 161), timeout=2, retries=3)
    szName_printer = await get_cmd(
                SnmpEngine(),
                CommunityData(community, mpModel=0),
                target,
                ContextData(),
                ObjectType(ObjectIdentity(test_oid))
        )
    szName_printer = szName_printer[3][0][1].asOctets()
    szName_printer = szName_printer.decode('utf-8')  # или 'utf-8', если версия

    print(szName_printer)

def Main():
    # locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    # d.decode('cp1251').encode('utf8')
    asyncio.run(GetTonersStatus())

if __name__ == "__main__":
    Main()


