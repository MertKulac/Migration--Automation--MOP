def layerTwoMop():

    backbone = '10.10.10.100'
    row = 1
    user = calisma.username
    pasw = calisma.password

    # kontrol noktalari
    bb_prompt = '.*one: '
    core_prompt = '.*>'
    list_prompt = "Type '*' to clear search or select one:"
    # Bekciye baglanti objesi
    ssh_bekci = paramiko.SSHClient()

    # Known host & RSA key
    ssh_bekci.load_system_host_keys()
    ssh_bekci.set_missing_host_key_policy(paramiko.AutoAddPolicy)

    # bekci baglantisi
    try:
        ssh_bekci.connect(hostname=backbone, username=user, password=pasw, port=2222)
        with SSHClientInteraction(ssh_bekci, timeout=10, display=False, buffer_size=65535) as core:
            try:
                core.expect(bb_prompt)
                core.send(hazirlanan_mop.popSwitch)
                core.expect([core_prompt, bb_prompt, list_prompt], timeout=20)
                if core.last_match == list_prompt:
                    core.send("1")
                    print("OK")
                while core.last_match == bb_prompt:
                    print('\n' + '#' * 70)
                    print('\nPop Switch ile baglanti saglanamadi.')
                    print('\n' + '#' * 70)
                    mevcut_UPE = input('Lutfen servisin calistigi upe ip adresini giriniz: ')
                    while not mevcut_UPE:
                        mevcut_UPE = input('Lutfen servisin calistigi upe ip adresini giriniz: ')
                    core.send('*')
                    core.expect(bb_prompt)
                    core.send(mevcut_UPE)
                    core.expect([core_prompt, bb_prompt], timeout=10)
                    if core.last_match == core_prompt:
                        break
                    else:
                        continue
                print('\n' + '#' * 70)
                print('PopSwitch\'e basariyla baglanildi! Kontroller saglaniyor.')
                print('\n' + '#' * 70)
                core.send('screen-length 0 temporary')
                core.expect(core_prompt)
                core.send('disp cur conf vsi vpls{}'.format(hazirlanan_mop.vlanID))
                core.expect(core_prompt)
                vsi_vpls = core.current_output_clean
                vsi_vpls = vsi_vpls.splitlines()
                vsi = [v for v in vsi_vpls if 'peer' in v]
                vsi = vsi[1].split()
                hazirlanan_mop.yedekUPE = vsi[1]

                vsi_vpls.remove("return")

                f = LOGfile.get_sheet(calisma.imt)
                style = xlwt.easyxf('pattern: pattern solid, fore_color custom ')
                styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom')
                for item in vsi_vpls:
                    f.write(row, 0, item, style)
                    f.write(row, 1, "", styleline)
                    row += 1

                core.send('disp vsi services vpls{}'.format((hazirlanan_mop.vlanID)))
                core.expect(core_prompt)
                layertwoInt = core.current_output_clean
                layertwoInt = layertwoInt.splitlines()
                layertwoInt = layertwoInt[5]
                layertwoInt = layertwoInt.split()
                layertwoInt = layertwoInt[0]
                core.send('disp cur int {}'.format(layertwoInt))
                core.expect(core_prompt)
                fizikselint = core.current_output_clean
                fizikselint = fizikselint.splitlines()
                fizikselint.remove("return")
                f = LOGfile.get_sheet(calisma.imt)
                style = xlwt.easyxf('pattern: pattern solid, fore_color  custom')
                styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom ')
                for r in fizikselint:
                    f.write(row, 0, r, style)
                    f.write(row, 1, "", styleline)
                    row += 1

                f = LOGfile.get_sheet(calisma.imt)
                k = f.row(0)
                stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                        'solid, fore_color custom; font:bold 1 ')
                k.write(0, "HI Old Config", stylebold)
                k.write(1, hazirlanan_mop.popSwitch, stylebold)

                core.send('disp vsi name vpls{} peer-info'.format(hazirlanan_mop.vlanID))
                core.expect(core_prompt)
                peerInfo = core.current_output_clean
                peerInfo = peerInfo.splitlines()

                postape(vsi_vpls, fizikselint)

            except Exception as e:
                print(e)
    except Exception as k:
        print('\n' + '#' * 70)


def kullanici():
    yeni_userlar = []
    for u in userlar:
        yeni_userlar.append(u.lower())
    return yeni_userlar

def ipbul(a):
    hazirlanan_mop.IP = []
    a = a.splitlines()
    for i in a:
        ip = re.search("ip\saddress", i)
        if ip:
            pp = ip.string
            hazirlanan_mop.IP.append(pp)
            #find ip

def preCheckArp(check):
    check = check.splitlines()
    stylebold = xlwt.easyxf('pattern: pattern solid, fore_color custom ; font:bold 1, height 260')
    f = LOGfile.get_sheet(calisma.imt)
    f.write(36, 0, "PRECHECK", stylebold)
    row = 39
    for item in check:
        f.write(row, 0, item)
        row += 1
    return row

def arpKontrol(arp):
    arp = arp.splitlines()
    arpBaslangic = 0
    arpbitis = 0
    sayac = 0
    while sayac < len(arp):
        if 'IP' in arp[sayac]:
            # print('arp icin baslangic satiri : {}'.format(sayac))
            arpBaslangic = sayac + 3
        if 'Total:' in arp[sayac]:
            # print('arp için bitiş satırı : {}'.format(sayac))
            arpbitis = sayac - 1
        sayac = sayac + 1
        #find arp

    arp = arp[arpBaslangic:arpbitis]
    yeniArp = []
    for i in arp:
        ip = i.split()
        if '.' in ip[0]:
            yeniArp.append(ip[0])

    return yeniArp


def bgpVrf(vpnins, col):
    vpnins = vpnins.splitlines()
    vpnins = vpnins[1:]
    bgp_vrf_final = []
    style = xlwt.easyxf('pattern: pattern solid, fore_color custom ')
    styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom ')

    for s in vpnins:
        if ' #' in s or "peer" in s:
            break
        else:
            bgp_vrf_final.append(s)
    f = LOGfile.get_sheet(calisma.imt)
    f.write(col + 1, 2, "bgp 34984", style)
    f.write(col + 1, 3, "", styleline)
    col = col + 2
    for b in bgp_vrf_final:
        f.write(col, 2, b, style)
        f.write(col, 3, "", styleline)
        col += 1
        #find bgpvrf


def vrfYazdir(vrfcikti, coll):
    vrfcikti = vrfcikti.splitlines()
    vrfcikti = vrfcikti[1:-2]

    style = xlwt.easyxf('pattern: pattern solid, fore_color custom ')
    styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom ')
    row = coll + 9
    f = LOGfile.get_sheet(calisma.imt)
    for v in vrfcikti:
        f.write(row, 2, v, style)
        f.write(row, 3, "", styleline)
        row = row + 1


def intYazdir(interface, x, vrf):
    # kontrol devrenin L2 - L3 olmasını ayiriyor. True l2, false l3
    oku = interface.splitlines()
    oku.remove("return")
    f = LOGfile.get_sheet(calisma.imt)
    styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom ')
    style = xlwt.easyxf('pattern: pattern solid, fore_color custom ')
    if x == 1:
        if len(vrf) > 0:
            row = 24
        else:
            row = 20
        for o in oku:
            f.write(row, 2, o, style)
            f.write(row, 3, "", styleline)
            row = row + 1

    elif x == 2:
        k = f.row(0)
        stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                'solid, fore_color custom; font:bold 1 ')
        k.write(2, "UPE Old Config", stylebold)
        k.write(3, calisma.oldUPE, stylebold)
        row2 = 1
        for o in oku:
            f.write(row2, 2, o, style)
            f.write(row2, 3, "", styleline)
            row2 = row2 + 1


def upeInterfaceAyir(intcikti):
    ciktibol = intcikti.splitlines()
    time.sleep(1)
    if len(ciktibol) > 11:
        ciktibol = ciktibol[-1].split(' ', 3)
        telekom = ciktibol[0]
        devre = ciktibol[0].split('E')

        vlanbul = ciktibol[0].split('.')
        vlanbul = vlanbul[1]
        hazirlanan_mop.vlanID = vlanbul
        devre = devre[1]

        # print('devrenin vlani : {}'.format(hazirlanan_mop.vlanID))
        if '/0/1.' in devre:
            # L3 devre bilgisi bulundu.
            hazirlanan_mop.layer_3_int = 'Virtual-Ethernet{}'.format(devre)
            # L2 devre bilgisi bul
            layer_2 = ''.join(devre)
            layer_2 = layer_2.replace('1.', '0.')
            hazirlanan_mop.layer_2_int = 'Virtual-Ethernet{}'.format(layer_2)
            # print('devrenin bulunan layer iki bilgisi : {}'.format(hazirlanan_mop.layer_2_int))
            solme = True
            return solme
        elif '/0/0.' in devre:
            # L2 devre bilgisi bulundu.
            hazirlanan_mop.layer_2_int = 'Virtual-Ethernet{}'.format(devre)
            # L3 devre bilgisi bul
            layer_3 = ''.join(devre)
            layer_3 = layer_3.replace('0.', '1.')
            hazirlanan_mop.layer_3_int = 'Virtual-Ethernet{}'.format(layer_3)
            # print('devrenin bulunan layer uc bilgisi : {}'.format(hazirlanan_mop.layer_3_int))
            solme = True
            return solme
        elif 'Global-VE' in devre:
            # L2 devre bilgisi bulundu.
            hazirlanan_mop.layer_2_int = 'Global-VE{}'.format(devre)
            # L3 devre bilgisi bul
            layer_3 = ''.join(devre)
            layer_3 = layer_3.replace('0.', '1.')
            hazirlanan_mop.layer_3_int = 'Global-VE{}'.format(layer_3)
            # print('devrenin bulunan layer uc bilgisi : {}'.format(hazirlanan_mop.layer_3_int))
            solme = True
            return solme
        elif '/1/1.' in devre:
            # L2 devre bilgisi bulundu.
            hazirlanan_mop.layer_2_int = 'Virtual-Ethernet{}'.format(devre)
            # L3 devre bilgisi bul
            layer_3 = ''.join(devre)

            layer_3 = layer_3.replace('0.', '1.')
            hazirlanan_mop.layer_3_int = 'Virtual-Ethernet{}'.format(layer_3)
            # print('devrenin bulunan layer uc bilgisi : {}'.format(hazirlanan_mop.layer_3_int))
            solme = True
            return solme
        elif '/1/2' in devre:
            # L3 devre bilgisi bulundu.
            hazirlanan_mop.layer_3_int = 'Virtual-Ethernet{}'.format(devre)
            # L2 devre bilgisi bul
            layer_2 = ''.join(devre)

            layer_2 = layer_2.replace('2.', '1.')

            hazirlanan_mop.layer_2_int = 'Virtual-Ethernet{}'.format(layer_2)
            # print('devrenin bulunan layer iki bilgisi : {}'.format(hazirlanan_mop.layer_2_int))
            solme = True
            return solme
        elif 'Eth' in telekom:
            # Servis telekom altyapisi kullaniyor
            hazirlanan_mop.layer_3_int = telekom
            hazirlanan_mop.layer_2_int = 'Telekom devresi oldugundan layer 2 devre bilgisi bulunmuyor.'
            solme = False
            return solme

    else:
        print('\n' + '#' * 70)
        print('\nDevre bulunamadi! IMT ID kontrol ederek tekrar deneyiniz!')
        print('\n' + '#' * 70)
        quit()


def bekciyeBaglan():
    # baglanti bilgileri

    backbone = '10.222.247.240'

    user = calisma.username
    pasw = calisma.password
    mevcut_UPE = calisma.oldUPE
    yeni_UPE = calisma.newUPE
    imtID = calisma.vlan
    vlanID = calisma.imt

    # kontrol noktalari
    bb_prompt = '.*one: '
    core_prompt = '.*>'

    # Bekciye baglanti objesi
    ssh_bekci = paramiko.SSHClient()

    # Known host & RSA key
    ssh_bekci.load_system_host_keys()
    ssh_bekci.set_missing_host_key_policy(paramiko.AutoAddPolicy)

    # bekci baglantisi
    try:
        ssh_bekci.connect(hostname=backbone, username=user, password=pasw, port=2222)
        print('\n' + '#' * 70)
        print('Bekciye basariyla baglanildi!')
        print('\n' + '#' * 70)
        with SSHClientInteraction(ssh_bekci, timeout=10, display=False, buffer_size=65535) as core:
            try:
                core.expect(bb_prompt)
                core.send(mevcut_UPE)
                core.expect([core_prompt, bb_prompt], timeout=20)
                while core.last_match == bb_prompt:
                    print('\n' + '#' * 70)
                    print('\nUPE ile baglanti saglanamadi.')
                    print('\n' + '#' * 70)
                    mevcut_UPE = input('Lutfen servisin calistigi upe ip adresini giriniz: ')
                    while not mevcut_UPE:
                        mevcut_UPE = input('Lutfen servisin calistigi upe ip adresini giriniz: ')
                    core.send('*')
                    core.expect(bb_prompt)
                    core.send(mevcut_UPE)
                    core.expect([core_prompt, bb_prompt], timeout=10)
                    if core.last_match == core_prompt:
                        break
                    else:
                        continue
                print('\n {} ip adresli upe baglantisi saglandi!'.format(mevcut_UPE))
                core.send('screen-length 0 temporary')
                core.expect(core_prompt)
                time.sleep(2)
                print('\n *************************************')
                print('\n {} numaralı vlan için precheck alınıyor'.format(vlanID))
                core.send('display interface description | i {}'.format(imtID))
                core.expect(core_prompt, timeout=5)
                ayirUPE = core.current_output_clean

                SOLME_or_TT = upeInterfaceAyir(ayirUPE)
                print(SOLME_or_TT)
                # true solme , false tt devresi
                time.sleep(1)
                core.send('display cur interface {}'.format(hazirlanan_mop.layer_3_int))
                core.expect(core_prompt, timeout=7)
                neforty_layerThree = core.current_output_clean
                print(neforty_layerThree)
                vrf = []
                intYazdir(neforty_layerThree, 2, vrf)
                postUpe(neforty_layerThree, 2)
                vrfNeforty = neforty_layerThree.splitlines()
                satir = len(vrfNeforty)
                # bgp_flag referans

                wifik = [b for b in vrfNeforty if 'wifi_ap' in b]

                vcon = [x for x in vrfNeforty if 'vpn-instance' in x]
                if len(vcon) > 0:
                    line = 56
                else:
                    line = 34

                styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom ')
                style = xlwt.easyxf('pattern: pattern solid, fore_color custom')
                if SOLME_or_TT == True:
                    time.sleep(2)
                    core.send('display cur interface {}'.format(hazirlanan_mop.layer_2_int))
                    core.expect(core_prompt, timeout=5)
                    neforty_layerTwo = core.current_output_clean
                    vsi_ilk = neforty_layerTwo
                    vsi_ilk = vsi_ilk.splitlines()
                    vsi_son = [x for x in vsi_ilk if 'l2vc' in x]
                    vsi_check = 1
                    satir2 = len(vsi_ilk)

                    # vsi l2 devre altındaysa buraya girecek , l2 vpn ise alttaki if'e girecek
                    if len(vsi_son) > 0:
                        vsi_son = vsi_son[0].split()
                        hazirlanan_mop.popSwitch = vsi_son[2]
                        vsi_check = 2

                    if vsi_check == 1:
                        core.send('disp cur conf vsi vpls{}'.format(hazirlanan_mop.vlanID))
                        core.expect(core_prompt, timeout=5)
                        vpls = core.current_output_clean
                        vpls = vpls.splitlines()
                        vsi_peer = [v for v in vpls if 'peer' in v]
                        vsi_peer = vsi_peer[0].split()
                        hazirlanan_mop.popSwitch = vsi_peer[1]
                        vsi_check = 3

                    intYazdir(neforty_layerTwo, 1, wifik)
                    postUpe(neforty_layerTwo, 1)

                    satir3 = satir2 + 3 + satir
                    if vsi_check == 3:

                        f = LOGfile.get_sheet(calisma.imt)

                        for v in vpls:
                            f.write(line, 2, v, style)
                            f.write(line, 3, "", styleline)
                            line = line + 1
                else:
                    satir3 = satir

                core.send('display cur interface {}'.format(hazirlanan_mop.layer_3_int))
                core.expect(core_prompt, timeout=7)
                intconf = core.current_output_clean
                ipbul(intconf)
                print(hazirlanan_mop.IP)

                ipadd = []

                for i in hazirlanan_mop.IP:
                    i = i.split()
                    ipadd.append(i[2])
                x = 0
                for m in range(len(ipadd)):
                    oktet = []
                    ipsplit = ipadd[x].split(".")
                    for j in ipsplit:
                        oktet.append(j)
                    last = int(oktet[3])

                    core.send("dis cur | inc {}.{}.{}.{}".format(oktet[0], oktet[1], oktet[2], last - 1))
                    core.expect(core_prompt, timeout=5)
                    prefix = core.current_output_clean
                    prefix = prefix.splitlines()
                    prefix = prefix[1:]
                    if len(prefix) > 0:
                        f = LOGfile.get_sheet(calisma.imt)
                        for s in prefix:
                            f.write(satir3, 2, s, style)
                            satir3 = satir3 + 1
                    x = x + 1

                core.send("dis cur configuration route-static | i {}".format(hazirlanan_mop.layer_3_int))
                core.expect(core_prompt, timeout=5)
                static = core.current_output_clean
                static = static.splitlines()
                static = static[1:]
                if len(static) > 0:
                    f = LOGfile.get_sheet(calisma.imt)
                    for s in static:
                        f.write(satir3, 2, s, style)
                        satir3 = satir3 + 1

                bgp_flag = None

                # vrf burada yakalaniyor

                # vrf varsa 2 yoksa 1
                vrf_kontrol = 1

                vrf = [x for x in vrfNeforty if 'vpn-instance' in x]
                if len(vrf) > 0:
                    vrf = vrf[0].split()
                    vrf = vrf[3]
                    vrf_kontrol = 2
                    # print(' Bulunan vrf: {}'.format(vrf))

                if vrf_kontrol == 2:
                    core.send('disp cur conf vpn-ins {}'.format(vrf))
                    core.expect(core_prompt, timeout=5)
                    vrfTanim = core.current_output_clean
                    vrfYazdir(vrfTanim, satir3)
                    core.send('disp cur conf bgp | begin vpn-instance {}'.format(vrf))
                    core.expect(core_prompt, timeout=5)
                    bgp_vrf = core.current_output_clean

                    bgpVrf(bgp_vrf, satir3)

                f = LOGfile.get_sheet(calisma.imt)
                bold = xlwt.easyxf(' font:bold 1 ')
                core.send('disp arp int {}'.format(hazirlanan_mop.layer_3_int))
                f.write(38, 0, 'display arp interface {}'.format(hazirlanan_mop.layer_3_int), bold)
                core.expect(core_prompt, timeout=5)
                arpcikti = core.current_output_clean
                preCheckArp(arpcikti)
                arpTablosu = arpKontrol(arpcikti)

                f = LOGfile.get_sheet(calisma.imt)

                for a in arpTablosu:
                    core.send('disp cur | i {}'.format(a))
                    core.expect(core_prompt, timeout=5)
                    routecheck = core.current_output_clean
                    routecheck = routecheck.splitlines()

                    if len(routecheck) >= 2:
                        routecheck = [x for x in routecheck if 'peer' in x]
                        if len(routecheck) >= 1:
                            peerCheck = routecheck[0].split()

                            # bgp flag devrede bgp olup olmadigini kontrol ediyor.
                            bgp_flag = peerCheck[0]
                            peerCheck = peerCheck[1]
                            rPolicy = [r for r in routecheck if 'route-policy' in r]
                            policy = rPolicy[0].split()
                            policyimport = policy[3]

                            core.send('disp cur conf route-policy {}'.format(policyimport))
                            core.expect(core_prompt, timeout=5)
                            alinanPolicy = core.current_output_clean
                            alinanPolicy = alinanPolicy.splitlines()
                            alinanPolicy.remove("return")

                            prefixbul = alinanPolicy[1]
                            prefixbul = prefixbul.split()
                            prefixbul = prefixbul[1]
                            prefixbul = prefixbul.split("-")
                            prefixbul = prefixbul[2]

                            if len(prefixbul) > 0:
                                # ifmatch = ifmatch[0].split()
                                # ifmatch = ifmatch[2]
                                core.send('disp cur | in ip ip-prefix | in {}'.format(prefixbul))
                                core.expect(core_prompt, timeout=5)
                                prefix = core.current_output_clean
                                prefix = prefix.splitlines()
                                prefix = prefix[1:]

                                for r in prefix:
                                    f.write(line, 2, r, style)
                                    f.write(line, 3, "", styleline)
                                    line = line + 1

                            line = line + 2
                            for r in alinanPolicy:
                                f.write(line, 2, r, style)
                                f.write(line, 3, "", styleline)
                                line = line + 1

                            line = line + 2
                            if a == peerCheck:
                                if vrf_kontrol == 1:
                                    f.write(line, 2, "bgp 34984", style)
                                    f.write(line, 3, "", style)
                                    line = line + 1
                                    for r in routecheck:
                                        f.write(line, 2, r, style)
                                        f.write(line, 3, "", styleline)
                                        line = line + 1
                                else:
                                    f.write(line, 2, "bgp 34984", style)
                                    f.write(line, 3, "", styleline)
                                    f.write(line + 1, 2, "ipv4-family vpn-instance {}".format(vrf), style)
                                    f.write(line + 1, 3, "", styleline)
                                    line = line + 2
                                    for r in routecheck:
                                        f.write(line, 2, r, style)
                                        f.write(line, 3, "", styleline)
                                        line = line + 1
                            else:
                                if vrf_kontrol == 1:
                                    f.write(line, 2, "bgp 34984")
                                    f.write(line, 3, "", styleline)
                                    line = line + 1
                                    for r in routecheck:
                                        f.write(line, 2, r, style)
                                        f.write(line, 3, "", styleline)
                                        line = line + 1
                                else:
                                    f.write(line, 2, "bgp 34984", style)
                                    f.write(line, 3, "", styleline)
                                    f.write(line + 1, 2, "ipv4-family vpn-instance {}".format(vrf), style)
                                    f.write(line + 1, 3, "", styleline)
                                    line = line + 2
                                    for r in routecheck:
                                        f.write(line, 2, r, style)
                                        f.write(line, 3, "", styleline)
                                        line = line + 1

                if bgp_flag == 'peer':

                    if vrf_kontrol == 1:
                        core.send('disp bgp peer {} verbose'.format(peerCheck))
                        core.expect(core_prompt, timeout=10)
                        bgpUpTime = core.current_output_clean
                        bgpUpTime = bgpUpTime.splitlines()
                        bgpEstablishTime = [b for b in bgpUpTime if 'Established,' in b]
                        line += 1
                        f.write(line, 2, bgpEstablishTime[0])
                        line += 1

                        core.send('disp bgp routing-table peer {} received-routes'.format(peerCheck))
                        core.expect(core_prompt, timeout=10)
                        alinanRoute = core.current_output_clean
                        alinanRoute = alinanRoute.splitlines()
                        f.write(line, 2, 'display bgp routing-table peer {} received-routes'.format(peerCheck, bold))
                        line += 1
                        for r in alinanRoute:
                            f.write(line, 2, r)
                            line = line + 1

                        core.send('disp bgp routi peer {} acc'.format(peerCheck))
                        core.expect(core_prompt, timeout=10)
                        kabuledilenRoute = core.current_output_clean
                        kabuledilenRoute = kabuledilenRoute.splitlines()
                        line = line + 2
                        f.write(line, 2, 'display bgp routing peer {} accepted'.format(peerCheck), bold)
                        line += 1
                        for r in kabuledilenRoute:
                            f.write(line, 2, r)
                            line = line + 1

                        core.send('disp bgp routin peer {} advert'.format(peerCheck))
                        core.expect(core_prompt, timeout=10)
                        anonslananRoute = core.current_output_clean
                        anonslananRoute = anonslananRoute.splitlines()

                        if len(anonslananRoute) < 20:
                            line = line + 2
                            f.write(line, 2, 'display bgp routing peer {} advert'.format(peerCheck), bold)
                            line += 1

                            for r in anonslananRoute:
                                f.write(line, 2, r)
                                line = line + 1

                        core.send('disp bgp peer {} verbose'.format(peerCheck))
                        core.expect(core_prompt, timeout=10)
                        bgpUpTime = core.current_output_clean
                        bgpUpTime = bgpUpTime.splitlines()
                        bgpEstablishTime = [b for b in bgpUpTime if 'Established,' in b]

                        f.write(line, 2, bgpEstablishTime[0])

                    if vrf_kontrol == 2:

                        core.send('disp bgp vpnv4 vpn-instance {} peer {} verbose'.format(vrf, peerCheck))
                        core.expect(core_prompt, timeout=10)
                        bgpUpTime = core.current_output_clean
                        bgpUpTime = bgpUpTime.splitlines()
                        bgpEstablishTime = [b for b in bgpUpTime if 'Established,' in b]
                        f.write(line, 2, 'disp bgp vpnv4 vpn-instance {} peer {} verbose'.format(vrf, peerCheck), bold)
                        line = line + 1
                        f.write(line, 2, bgpEstablishTime[0])
                        line = line + 2

                        core.send('disp bgp vpnv4 vpn-in {} routi peer {} rece'.format(vrf, peerCheck))
                        core.expect(core_prompt, timeout=10)
                        alinanRoute = core.current_output_clean
                        alinanRoute = alinanRoute.splitlines()
                        line = line + 2
                        f.write(line, 2,
                                'disp bgp vpnv4 vpn-instance {} routing peer {} received'.format(vrf, peerCheck), bold)
                        line += 1
                        for r in alinanRoute:
                            f.write(line, 2, r)
                            line = line + 1

                        core.send('disp bgp vpnv4 vpn-in {} routi peer {} acc'.format(vrf, peerCheck))
                        core.expect(core_prompt, timeout=10)
                        kabuledilenRoute = core.current_output_clean
                        kabuledilenRoute = kabuledilenRoute.splitlines()
                        line = line + 2
                        f.write(line, 2, 'disp bgp vpnv4 vpn-instance {} routing peer {} acc'.format(vrf, peerCheck),
                                bold)
                        line += 1
                        for r in kabuledilenRoute:
                            f.write(line, 2, r)
                            line = line + 1

                        core.send('disp bgp vpnv4 vpn-in {} routi peer {} adv'.format(vrf, peerCheck))
                        core.expect(core_prompt, timeout=10)
                        anonslananRoute = core.current_output_clean
                        anonslananRoute = anonslananRoute.splitlines()
                        line = line + 2
                        f.write(line, 2, 'disp bgp vpnv4 vpn-instance {} routing peer {} adv'.format(vrf, peerCheck),
                                bold)
                        line += 1

                        for r in anonslananRoute:
                            f.write(line, 2, r)
                            line = line + 1

                k = 56

                if vrf_kontrol == 2:

                    for p in arpTablosu:
                        core.send('ping -vpn-instance {} {}'.format(vrf, p))
                        core.expect(core_prompt, timeout=20)
                        pingcikti = core.current_output_clean
                        pingcikti = pingcikti.splitlines()
                        for r in pingcikti:
                            f.write(k, 0, r)
                            k = k + 1

                else:
                    for t in arpTablosu:
                        core.send('ping {}'.format(t))
                        core.expect(core_prompt, timeout=20)
                        intping = core.current_output_clean
                        intping = intping.splitlines()
                        for r in intping:
                            f.write(k, 0, r)
                            k = k + 1



            except Exception as e:
                print(e)
    except Exception as e:
        print(
            'Bekciye baglanti saglanamadi! \nSirket agina bagli oldugunuzdan ve giris bilgilerinin dogru oldugundan emin olunuz!')
        print('\n' + '#' * 70)
        print(e)


def postape(vsi_vpls, fizikselint):
    yesil = []
    f = LOGfile.get_sheet(calisma.imt)

    style = xlwt.easyxf('pattern: pattern solid, fore_color custom2 ')
    styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom2')

    print("OK")

    stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '

                            'solid, fore_color custom2; font:bold 1 ')

    if peer.degisim == True and len(peer.new2) > 0:
        k = f.row(0)

        k.write(4, "APE New Config", stylebold)
        k.write(5, "X.X.X.X", stylebold)

        vsi_id = [b for b in vsi_vpls if 'vsi-id' in b]

        # vsi_id = vsi_id[1]

        desc = [b for b in vsi_vpls if 'description' in b]

        # desc = desc[1]

        f.write(1, 4, "#", style)
        f.write(2, 4, "vsi vpls{} static".format(calisma.imt), style)
        f.write(3, 4, desc, style)
        f.write(4, 4, "mac-withdraw enable", style)
        f.write(5, 4, "interface-status-change mac-withdraw enable", style)
        f.write(6, 4, "mac-limit maximum 1000", style)
        f.write(7, 4, "pwsignal ldp", style)
        f.write(8, 4, vsi_id, style)
        f.write(9, 4, "peer {} tnl-policy lsp-lb".format(peer.new1), style)
        f.write(10, 4, "peer {} tnl-policy lsp-lb".format(peer.new2), style)
        f.write(11, 4, "protect-group vlan{}".format(calisma.imt), style)
        f.write(12, 4, "protect-mode pw-redundancy master", style)
        f.write(13, 4, "reroute immediately", style)
        f.write(14, 4, "holdoff 1", style)
        f.write(15, 4, "peer {} preference 1".format(peer.new1), style)
        f.write(16, 4, "peer {} preference 2".format(peer.new2), style)
        f.write(17, 4, "traffic-statistics enable", style)
        f.write(18, 4, "mtu 9000", style)
        f.write(19, 4, "#", style)

        for i in range(18):
            f.write(i + 1, 5, "", styleline)

        fizikselint[1] = "interface zzz"

        row2 = 21
        for i in fizikselint:
            f.write(row2, 4, i, style)
            f.write(row2, 5, "", styleline)
            row2 = row2 + 1

    elif peer.degisim == True and len(peer.new2) == 0:

        k = f.row(0)

        k.write(4, "APE New Config", stylebold)
        k.write(5, "X.X.X.X", stylebold)

        vsi_id = [b for b in vsi_vpls if 'vsi-id' in b]

        # vsi_id = vsi_id[1]

        desc = [b for b in vsi_vpls if 'description' in b]

        # desc = desc[1]

        f.write(1, 4, "#", style)
        f.write(2, 4, "vsi vpls{} static".format(calisma.imt), style)
        f.write(3, 4, desc, style)
        f.write(4, 4, "mac-withdraw enable", style)
        f.write(5, 4, "interface-status-change mac-withdraw enable", style)
        f.write(6, 4, "mac-limit maximum 1000", style)
        f.write(7, 4, "pwsignal ldp", style)
        f.write(8, 4, vsi_id, style)
        f.write(9, 4, "peer {} tnl-policy lsp-lb".format(peer.new1), style)
        f.write(10, 4, "peer {} tnl-policy lsp-lb".format(hazirlanan_mop.yedekUPE), style)
        f.write(11, 4, "protect-group vlan{}".format(calisma.imt), style)
        f.write(12, 4, "protect-mode pw-redundancy master", style)
        f.write(13, 4, "reroute immediately", style)
        f.write(14, 4, "holdoff 1", style)
        f.write(15, 4, "peer {} preference 1".format(peer.new1), style)
        f.write(16, 4, "peer {} preference 2".format(hazirlanan_mop.yedekUPE), style)
        f.write(17, 4, "traffic-statistics enable", style)
        f.write(18, 4, "mtu 9000", style)
        f.write(19, 4, "#", style)

        for i in range(18):
            f.write(i + 1, 5, "", styleline)

        fizikselint[1] = "interface zzz"

        row2 = 21
        for i in fizikselint:
            f.write(row2, 4, i, style)
            f.write(row2, 5, "", styleline)
            row2 = row2 + 1


    elif len(peer.new2) > 0 and peer.degisim == False:
        k = f.row(0)

        k.write(4, "APE New Config", stylebold)
        k.write(5, hazirlanan_mop.popSwitch, stylebold)
        f.write(2, 4, "vsi vpls{} static".format(calisma.imt), style)
        f.write(3, 4, "pwsignal ldp", style)
        f.write(4, 4, "undo peer {}".format(calisma.oldUPE), style)
        f.write(5, 4, "undo peer {}".format(hazirlanan_mop.yedekUPE), style)
        f.write(6, 4, "peer {} tnl-policy lsp-lb".format(peer.new1), style)
        f.write(7, 4, "peer {} tnl-policy lsp-lb".format(peer.new2), style)
        f.write(8, 4, "protect-group vlan{}".format(calisma.imt), style)
        f.write(9, 4, "protect-mode pw-redundancy master", style)
        f.write(10, 4, "peer {} preference 1".format(peer.new1), style)
        f.write(11, 4, "peer {} preference 2".format(peer.new2), style)

        for i in range(10):
            f.write(i + 2, 5, "", styleline)

    elif len(peer.new2) == 0 and peer.degisim == False:

        k = f.row(0)

        k.write(4, "APE New Config", stylebold)
        k.write(5, hazirlanan_mop.popSwitch, stylebold)
        f.write(2, 4, "vsi vpls{} static".format(calisma.imt), style)
        f.write(3, 4, "pwsignal ldp", style)
        f.write(4, 4, "undo peer {}".format(calisma.oldUPE), style)
        f.write(5, 4, "peer {} tnl-policy lsp-lb".format(peer.new1), style)
        f.write(6, 4, "protect-group vlan{}".format(calisma.imt), style)
        f.write(7, 4, "protect-mode pw-redundancy master", style)
        f.write(8, 4, "peer {} preference 1".format(peer.new1), style)

        for i in range(7):
            f.write(i + 2, 5, "", styleline)

    pass


def postUpe(interface, x):
    oku = interface.splitlines()
    oku.remove("return")
    f = LOGfile.get_sheet(calisma.imt)
    styleline = xlwt.easyxf('border: right thick; pattern: pattern solid, fore_color custom2 ')
    style = xlwt.easyxf('pattern: pattern solid, fore_color custom2 ')
    if x == 1 and peer.degisim and len(peer.new2) > 0:
        oku[7] = "mpls l2vc x.x.x.x tunnel-policy lsp-lb"
        row = 21
        for o in oku:
            f.write(row, 6, o, style)
            f.write(row, 7, "", styleline)
            row = row + 1
        desc = [v for v in oku if "description" in v]
        ind = oku.index(desc[0])

        oku[ind] = oku[ind].replace("01", "02")
        rowyedek = 21
        for o in oku:
            f.write(rowyedek, 8, o, style)
            f.write(rowyedek, 9, "", styleline)
            rowyedek = rowyedek + 1
        k = f.row(0)
        stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                'solid, fore_color custom2; font:bold 1 ')
        k.write(8, "Slave UPE New Config", stylebold)
        k.write(9, "{}".format(peer.new2), stylebold)
    elif x == 1 and peer.degisim == False and len(peer.new2) > 0:
        row = 21
        for o in oku:
            f.write(row, 6, o, style)
            f.write(row, 7, "", styleline)
            row = row + 1

        desc = [v for v in oku if "description" in v]
        ind = oku.index(desc[0])

        oku[ind] = oku[ind].replace("01", "02")
        rowyedek = 21
        for o in oku:
            f.write(rowyedek, 8, o, style)
            f.write(rowyedek, 9, "", styleline)
            rowyedek = rowyedek + 1
        k = f.row(0)
        stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                'solid, fore_color custom2; font:bold 1 ')
        k.write(8, "Slave UPE New Config", stylebold)
        k.write(9, "{}".format(peer.new2), stylebold)


    elif x == 1 and peer.degisim == False and len(peer.new2) == 0:
        row = 21
        for o in oku:
            f.write(row, 6, o, style)
            f.write(row, 7, "", styleline)
            row = row + 1
    elif x == 1 and peer.degisim and len(peer.new2) == 0:
        oku[7] = "mpls l2vc x.x.x.x tunnel-policy lsp-lb"
        row = 21
        for o in oku:
            f.write(row, 6, o, style)
            f.write(row, 7, "", styleline)
            row = row + 1

    elif x == 2 and len(peer.new2) == 0:
        k = f.row(0)
        stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                'solid, fore_color custom2; font:bold 1 ')
        k.write(6, "UPE New Config", stylebold)
        k.write(7, "{}".format(peer.new1), stylebold)
        row2 = 1
        for o in oku:
            f.write(row2, 6, o, style)
            f.write(row2, 7, "", styleline)
            row2 = row2 + 1

    elif x == 2 and len(peer.new2) > 0:
        k = f.row(0)
        stylebold = xlwt.easyxf('border: top thick, right thick, bottom thick, left thick; pattern: pattern '
                                'solid, fore_color custom2; font:bold 1 ')
        k.write(6, "UPE New Config", stylebold)
        k.write(7, "{}".format(peer.new1), stylebold)
        row2 = 1
        for o in oku:
            f.write(row2, 6, o, style)
            f.write(row2, 7, "", styleline)
            row2 = row2 + 1
        desc = [v for v in oku if "description" in v]
        ind = oku.index(desc[0])
        row2yedek = 1
        oku[ind] = oku[ind].replace("01", "02")

        for o in oku:
            f.write(row2yedek, 8, o, style)
            f.write(row2yedek, 9, "", styleline)
            row2yedek = row2yedek + 1

    pass


class servis(object):

    def __init__(self, vlanID, anaUPE, yedekUPE, popSwitch, layer_2_int, ne40_l2_cikti, layer_3_int, ne40_l3_cikti,
                 routeType, IP):
        self.vlanID = vlanID
        self.anaUPE = anaUPE
        self.yedekUPE = yedekUPE
        self.popSwitch = popSwitch
        self.layer_2_int = layer_2_int
        self.ne_40_l2_cikti = ne40_l2_cikti
        self.layer_3_int = layer_3_int
        self.ne40_l3_cikti = ne40_l3_cikti
        self.routeType = routeType
        self.IP = IP


if __name__ == '__main__':
    from PyQt5 import QtWidgets

    LOGfile = xlwt.Workbook("test.xlsx")
    xlwt.add_palette_colour("custom", 0x21)
    LOGfile.set_colour_RGB(0x21, 255, 188, 3)
    xlwt.add_palette_colour("custom2", 0x22)
    LOGfile.set_colour_RGB(0x22, 0, 195, 17)


    class mop(object):

        def __init__(self, username, password, oldUPE, newUPE, vlan):
            self.username = username
            self.password = password
            self.oldUPE = oldUPE
            self.newUPE = newUPE
            self.vlan = vlan


    class peer(object):
        def __init__(self, new1, new2, degisim):
            self.new1 = new1
            self.new2 = new2
            self.degisim = degisim


    hazirlanan_mop = servis(None, None, None, None, None, None, None, None, None, None)
    calisma = mop(None, None, None, None, None)


    # hazirlanan_mop = servis(None, None, None, None, None, None, None, None, None)

    class Pencere(QtWidgets.QWidget):

        def __init__(self):
            super().__init__()

            self.init_ui()

        def init_ui(self):
            self.u = QtWidgets.QLabel("USERNAME")
            self.p = QtWidgets.QLabel("PASSWORD")
            self.up = QtWidgets.QLabel("OLD UPE")
            self.d = QtWidgets.QLabel("IMT:VLAN")
            self.Ape = QtWidgets.QCheckBox("APE değişecek mi? Evet için tıklayın")
            self.o = QtWidgets.QLabel("NEW MASTER UPE")
            self.o2 = QtWidgets.QLabel("NEW SLAVE UPE")
            self.old1 = QtWidgets.QLineEdit()
            self.old2 = QtWidgets.QLineEdit()
            self.user = QtWidgets.QLineEdit()
            self.password = QtWidgets.QLineEdit()
            self.password.setEchoMode(QtWidgets.QLineEdit.Password)
            self.imt = QTextEdit()
            self.buton = QtWidgets.QPushButton("precheck")
            self.upe = QtWidgets.QLineEdit()
            self.check = QtWidgets.QCheckBox("TT")

            v_box = QtWidgets.QVBoxLayout()
            v_box.addWidget(self.u)
            v_box.addWidget(self.user)
            v_box.addWidget(self.p)
            v_box.addWidget(self.password)
            v_box.addWidget(self.up)
            v_box.addWidget(self.upe)
            v_box.addWidget(self.Ape)
            v_box.addWidget(self.o)
            v_box.addWidget(self.old1)
            v_box.addWidget(self.o2)
            v_box.addWidget(self.old2)
            v_box.addWidget(self.d)
            v_box.addStretch()
            v_box.addWidget(self.imt)
            v_box.addWidget(self.buton)
            v_box.addWidget(self.check)

            h_box = QtWidgets.QHBoxLayout()
            h_box.addStretch()
            h_box.addLayout(v_box)
            h_box.addStretch()

            self.setLayout(h_box)
            self.setWindowTitle("PRECHECK MOP")
            self.setGeometry(50, 50, 400, 500)

            self.show()

            self.buton.clicked.connect(self.click)

        def click(self):
            sender = self.sender()

            if sender.text() == "precheck":
                kullanici = self.user.text()
                sifre = self.password.text()
                devre = self.imt.toPlainText()
                upe = self.upe.text()
                peer1 = self.old1.text()
                peer2 = self.old2.text()

                devre = devre.splitlines()

                for i in range(len(devre)):

                    devre2 = devre[i].split(":")
                    imt = devre2[0]
                    vlan = devre2[1]

                    calisma.username = kullanici
                    calisma.password = sifre
                    calisma.oldUPE = upe
                    calisma.vlan = imt
                    calisma.imt = vlan

                    peer.new1 = peer1
                    peer.new2 = peer2
                    peer.degisim = self.Ape.isChecked()
                    LOGfile.add_sheet(vlan)
                    bekciyeBaglan()

                    if (self.check.isChecked()):
                        continue
                    else:
                        layerTwoMop()

                    time.sleep(3)

            self.close()


    app = QtWidgets.QApplication(sys.argv)

    pencere = Pencere()

    app.exec_()

    LOGfile.save("deneme.xls")

sys.exit()
