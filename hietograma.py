import numpy as np
from xlwings import Workbook, Range, Chart

def hietograma():

    def lluvia_max():
        muestra = Range('LLUVIA','C5').table.value
        #numero de datos
        DataNr = len(muestra)
        #cambiamos tipo de dato a 'array'
        muestra = np.asarray(muestra)
        #todos valores de los parametrod de Gumbel hasta 58 datos
        Yn = Range('+','C4:C29').value
        Sigma = Range('+','D4:D29').value
        Indeks = Range('+','B4:B29').value.index(DataNr)
        
        #Valores adoptados por los parametros de la funcion Gumbel:
        Yn = Yn[Indeks]
        Sigma = Sigma[Indeks]
        ### empty lists creation  
        s=[]
        m=[]
        alfa=[]
        mu=[]
        P=[]
        Prob=[]
        #los parametros de la funcion de Gumbel para cada estacion
        #ddof: "Delta Degrees of Freedom" by default 0
        for i in range(0,4):
            columna=muestra[:, i]
            s.append(np.std(columna,ddof=1)) 
            m.append(np.mean(columna))
            alfa.append(s[i]/Sigma)
            mu.append(m[i]-Yn*alfa[i])
            P.append(mu[i]-alfa[i]*np.log(np.log(T/(T-1))))
            Prob.append(1-1/T)
        ### PESOS
        dist=np.asarray([3968, 14046, 13251, 12596]) #distancias...
        pesos=1/dist**2
        sq=[1/i**2 for i in dist]
        pesos=pesos/sum(sq)
        ### Pmax24
        return sum(P*pesos)

    wb = Workbook.caller()
    
    #### clear table
    Range('R5').table.clear_contents()

    #### introduccion de los datos
    ####periodo de  retorno
    T =int(Range('INTERFACE', 'D8').value) 
    #### duracion de lluvia
    dur =int(Range('INTERFACE', 'D10').value)
    #### intervalo de tiempo
    inval= int(Range('INTERFACE', 'D12').value)
    #### En este script I1/Id se denomina X
    X = int(Range('INTERFACE', 'D14').value)

    #### Precipitaciones maximas diarias
    P24=Range('P24', 'C14:J14').value

    if Range('INTERFACE', 'D12').value=='Ya tengo P24':
        P24=Range('INTERFACE', 'D18').value
    else:
        #P24=Range('P24', 'C14:J14').value
        P24=lluvia_max()
    print(P24)
    #### periodos de retorno disponibles
    Tr = [2, 5, 10, 25, 50, 100, 500, 1000]
    ind= Tr.index(T)

    #### generacion del tiempo eje X
    time =np.arange(inval,dur*60+inval,inval)
    Range('INTERFACE','R5').value = np.arange(inval,dur*60+inval,inval)[:, np.newaxis]


    #### Precipitaciones maximas diarias mm/hr
    #Id=[i/24 for i in P24]
    Id=P24/24

    #### Intensidad de lluvia mm/hr
    I=Id*(X)**((28**0.1-(time/60)**0.1)/(28**0.1-1))
    Range('INTERFACE','S5').value = I[:, np.newaxis]

    lluvia = I*time/60
    Range('INTERFACE','T5').value = lluvia[:, np.newaxis]
    x=lluvia[0]
    #### Incremental depth 
    IncDep =[]
    for i in range(1,len(lluvia)):
        IncDep.append(lluvia[i]-lluvia[i-1])
    IncDep.insert(0,x)
    IncDep = np.asarray(IncDep)
    Range('INTERFACE','U5').value = IncDep[:, np.newaxis]

    #### Ordenacion methodo de bloques alternos
    nr=Range('U5').table.last_cell.row
    tt = sorted(Range('INTERFACE','U5:U'+str(nr)).value)
    new_tt=sorted(tt)[1::2]+sorted(tt)[::2][::-1]
    new_tt = np.asarray(new_tt)
    Range('INTERFACE','V5:V'+str(nr)).value=new_tt[:, np.newaxis]

    #### Chart 
    chart = Chart('Wykres 4',source_data=Range('V5:V'+str(nr)).table)
