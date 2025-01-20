#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

#############################
# Fichiers et paramètres

    #Fichiers
    file_cas = ("Données/Donnees_etude_de_cas_ETE305.xlsx") # données de l'étude de cas
    file_tp = ("Données/TP O&D.xlsx") # données du TP OED
    num_semaine = ("Données/week.xlsx") # indice de la semaine modélisée, donné par le script python
    effet_bord = ("Données/bord.xlsx") # données associées à l'effet de bord

    #Paramètres
    week = XLSX.readdata(num_semaine, "Sheet1", "A2")
    week = Int64(week[1])
    Tmax = 7 * 24 # simulation pour une semaine, heure par heure
    delimitation_min = Tmax * (week)
    delimitation_max = Tmax * (week+1)

#############################

#############################
# Données historiques pour énergies fatales et conso

    ## Consommation (MW) heure par heure de juillet à juin sur un an
    conso = XLSX.readdata(file_tp, "TP", "C" * string(10 + delimitation_min) * ":C" * string(10 + delimitation_max))

    ## Production (MW) Heure par heure de juillet à juin sur un an
    hydro_fatal = XLSX.readdata(file_cas, "Détails historique hydro", "M" * string(2 + delimitation_min) * ":M" * string(2 + delimitation_max)) # Production hydraulique fatale
    lacs = XLSX.readdata(file_cas, "Détails historique hydro", "N" * string(2 + delimitation_min) * ":N" * string(2 + delimitation_max)) # Production hydraulique issue des lacs
    step = XLSX.readdata(file_cas, "Détails historique hydro", "O" * string(2 + delimitation_min) * ":O" * string(2 + delimitation_max)) # Production hydraulique issue des STEP
    wind = XLSX.readdata(file_tp, "TP", "D" * string(10 + delimitation_min) * ":D" * string(10 + delimitation_max)) # Production éolienne
    solar = XLSX.readdata(file_tp, "TP", "E" * string(10 + delimitation_min) * ":E" * string(10 + delimitation_max)) # Production solaire
    thermal_fatal = XLSX.readdata(file_tp, "TP", "G" * string(10 + delimitation_min) * ":G" * string(10 + delimitation_max)) # Production thermique fatale

    # Production résiduelle
    Pres = wind + solar + hydro_fatal + thermal_fatal

#############################

#############################
#Définition du parc
    #Clusters : Nucléaire, Gaz, Charbon, fuel, Cogénération, Hydro


    ## Données du nucléaire
    Nnuc = 3 # nombre d'unités de génération
    names_nuc = XLSX.readdata(file_cas, "Parc électrique", "A2:A4") # noms des centrales 
    dict_nuc = Dict(i => names_nuc[i] for i in 1:Nnuc)
    costs_nuc = XLSX.readdata(file_cas, "Parc électrique", "H2:H4") # coûts de production
    Pmin_nuc = XLSX.readdata(file_cas, "Parc électrique", "F2:F4") # puissance minimale (MW)
    Pmax_nuc = XLSX.readdata(file_cas, "Parc électrique", "E2:E4") # puissance maximale (MW)
    dmin_nuc = XLSX.readdata(file_cas, "Parc électrique", "G2:G4") # intervalle minimal d'utilisation (heures)

    # # Gestion des effets de bord
    # if week == 1
    #     for g in 1:Nnuc
    #         UPnuc[1,g]==0
    #         DOnuc[1,g]==0
    #     end
    # else
    #     for g in 1:Nnuc
    #         UPnuc[1,g] = XLSX.readdata(effet_bord, "Sheet1", "A2:D2")
    #         DO_nuc[1,g] = XLSX.readdata(effet_bord, "Sheet1", "B2:D2")
    #     end


    ## Données du gaz
    Ngaz = 6 # nombre d'unités de génération
    names_gaz = XLSX.readdata(file_cas, "Parc électrique", "A5:A10") # noms des centrales 
    dict_gaz = Dict(i => names_gaz[i] for i in 1:Ngaz)
    costs_gaz = XLSX.readdata(file_cas, "Parc électrique", "H5:H10") # coûts de production
    Pmin_gaz = XLSX.readdata(file_cas, "Parc électrique", "F5:F10") # puissance minimale (MW)
    Pmax_gaz = XLSX.readdata(file_cas, "Parc électrique", "E5:E10") # puissance maximale (MW)
    dmin_gaz = XLSX.readdata(file_cas, "Parc électrique", "H5:H10") # intervalle minimal d'utilisation (heures)

    ## Données du charbon
    Ncoal = 3 # nombre d'unités de génération
    names_coal = XLSX.readdata(file_cas, "Parc électrique", "A12:A14") # noms des centrales 
    dict_coal = Dict(i => names_coal[i] for i in 1:Ncoal)
    costs_coal = XLSX.readdata(file_cas, "Parc électrique", "H12:H14") # coûts de production
    Pmin_coal = XLSX.readdata(file_cas, "Parc électrique", "F12:F14") # puissance minimale (MW)
    Pmax_coal = XLSX.readdata(file_cas, "Parc électrique", "E12:E14") # puissance maximale (MW)
    dmin_coal = XLSX.readdata(file_cas, "Parc électrique", "H12:H14") # intervalle minimal d'utilisation (heures)

    ## Données du fioul
    Nfuel = 2 # nombre d'unités de génération
    names_fuel = XLSX.readdata(file_cas, "Parc électrique", "A17:A18") # noms des centrales 
    dict_fuel = Dict(i => names_fuel[i] for i in 1:Nfuel)
    costs_fuel = XLSX.readdata(file_cas, "Parc électrique", "H17:H18") # coûts de production
    Pmin_fuel = XLSX.readdata(file_cas, "Parc électrique", "F17:F18") # puissance minimale (MW)
    Pmax_fuel = XLSX.readdata(file_cas, "Parc électrique", "E17:E18") # puissance maximale (MW)
    dmin_fuel = XLSX.readdata(file_cas, "Parc électrique", "H17:H18") # intervalle minimal d'utilisation (heures)

    ## Données de la cogénération
    Ncogen = 1 # nombre d'unités de génération
    names_cogen = XLSX.readdata(file_cas, "Parc électrique", "A11") # noms des centrales 
    dict_cogen = Dict(i => names_cogen[i] for i in 1:Ncogen)
    costs_cogen = XLSX.readdata(file_cas, "Parc électrique", "H11") * ones(Ncogen)  # coûts de production
    Pmin_cogen = zeros(Ncogen) # puissance minimale (MW) (nulle ici)
    Pmax_cogen = XLSX.readdata(file_cas, "Parc électrique", "C11") * ones(Ncogen) # puissance maximale (MW)

    ## Données du réservoir hydraulique
    Nhy = 1 # nombre d'unités de génération
    Pmin_hy = zeros(Nhy) # puissance minimale (MW) (nulle ici)
    Pmax_hy = XLSX.readdata(file_cas, "Parc électrique", "C20") * ones(Nhy) # puissance maximale (MW)
    # RAJOUTER LE STOCK HYDRO À CHAQUE INSTANT
    #e_hy_stock = XLSX.readdata(file_cas, "Stock hydro", "B4")/100 * XLSX.readdata(file_cas, "Stock hydro", "B1")*1000000 * ones(Nhy) #MWh, le stock hydro, B4 = % du stock B1 en TWh
    e_hy_stockmax = XLSX.readdata(file_cas, "Stock hydro", "D1") * 1000 * ones(Nhy) #MWh, le stock hydro max du barrage
    e_hy = 1000 * ones(Nhy) # MWh, au début (juste pour faire run pour l'instant)
    e_hy_low = e_hy_stockmax .* XLSX.readdata(file_cas, "Stock hydro", "D" * string(4 + delimitation_min) * ":D" * string(4 + delimitation_max)) / 100 # plus bas niveau du réservoir au premier jour
    e_hy_high = e_hy_stockmax .* XLSX.readdata(file_cas, "Stock hydro", "E" * string(4 + delimitation_min) * ":E" * string(4 + delimitation_max)) / 100 # plus haut niveau du parc du réservoir au premier jour

    costs_hy = 0 * ones(Nhy) #MWh, gratuit

    # Apports en eau de juillet à juin en MWh, heure par heure (hypothèse : pluie continue égale à la moyenne mensuelle)
    apports_hydro = XLSX.readdata(file_cas, "Stock hydro", "F" * string(4 + delimitation_min) * ":F" * string(4 + delimitation_max))

    #data for STEP
    Pmax_STEP = XLSX.readdata(file_cas, "Parc électrique", "E21") #MW
    rSTEP = 0.75 # Le rendement de la STEP
    e_STEP = XLSX.readdata(file_cas, "Parc électrique", "C21") # POURQUOI ON A MODÉLISÉ LA CAPA JUSTE POUR STEP DÉJÀ ?

    #costs vectors
    cnuc = repeat(costs_nuc', Tmax) #cost of nuclear generation €/MWh
    cgaz = repeat(costs_gaz', Tmax) #cost of gaz generation €/MWh
    ccoal = repeat(costs_coal', Tmax) #cost of coal generation €/MWh
    cfuel = repeat(costs_fuel', Tmax) #cost of fuel generation €/MWh
    ccogen = repeat(costs_cogen', Tmax) #cost of cogen generation €/Mwh
    chy = repeat(costs_hy', Tmax) #cost of hydro generation €/MWh
    cuns = 5000 * ones(Tmax) #cost of unsupplied energy €/MWh # La défaillance. On met un coût très élevé.
    cexc = 0 * ones(Tmax) #cost of in excess energy €/MWh

############################# 

#############################
#Création du modèle d'optimisation

    model = Model(HiGHS.Optimizer)
    set_silent(model) # Pour ne pas afficher de texte

#############################

#############################
#Définition des variables

    #Nuclear generation variables

    @variable(model, Pnuc[1:Tmax, 1:Nnuc] >= 0)
    @variable(model, UCnuc[1:Tmax, 1:Nnuc], Bin)
    @variable(model, UPnuc[1:Tmax, 1:Nnuc], Bin)
    @variable(model, DOnuc[1:Tmax, 1:Nnuc], Bin)
    @variable(model, timeUPnuc[1:Tmax, 1:Nnuc] >= 0)
    @variable(model, timeDOnuc[1:Tmax, 1:Nnuc] >= 0)

    #gaz generation variables

    @variable(model, Pgaz[1:Tmax, 1:Ngaz] >= 0)
    @variable(model, UCgaz[1:Tmax, 1:Ngaz], Bin)# Unit Commitment, ie =1 si allumé, 0 sinon
    @variable(model, UPgaz[1:Tmax, 1:Ngaz], Bin)# =1 quand s'allume
    @variable(model, DOgaz[1:Tmax, 1:Ngaz], Bin)# =1 quand s'éteint
    @variable(model, timeUPgaz[1:Tmax, 1:Ngaz] >= 0)
    @variable(model, timeDOgaz[1:Tmax, 1:Ngaz] >= 0)

    #coal generation variables

    @variable(model, Pcoal[1:Tmax, 1:Ncoal] >= 0)
    @variable(model, UCcoal[1:Tmax, 1:Ncoal], Bin)
    @variable(model, UPcoal[1:Tmax, 1:Ncoal], Bin)
    @variable(model, DOcoal[1:Tmax, 1:Ncoal], Bin)
    @variable(model, timeUPcoal[1:Tmax, 1:Ncoal] >= 0)
    @variable(model, timeDOcoal[1:Tmax, 1:Ncoal] >= 0)

    #fuel generation variables

    @variable(model, Pfuel[1:Tmax, 1:Nfuel] >= 0)
    @variable(model, UCfuel[1:Tmax, 1:Nfuel], Bin)
    @variable(model, UPfuel[1:Tmax, 1:Nfuel], Bin)
    @variable(model, DOfuel[1:Tmax, 1:Nfuel], Bin)
    @variable(model, timeUPfuel[1:Tmax, 1:Nfuel] >= 0)
    @variable(model, timeDOfuel[1:Tmax, 1:Nfuel] >= 0)

    #cogen generation variables

    @variable(model, Pcogen[1:Tmax, 1:Ncogen] >= 0)
    @variable(model, UCcogen[1:Tmax, 1:Ncogen], Bin)
    @variable(model, UPcogen[1:Tmax, 1:Ncogen], Bin)
    @variable(model, DOcogen[1:Tmax, 1:Ncogen], Bin)

    #hydro generation variables
    @variable(model, Phy[1:Tmax, 1:Nhy] >= 0)

    #unsupplied energy variables
    @variable(model, Puns[1:Tmax] >= 0)

    #in excess energy variables
    @variable(model, Pexc[1:Tmax] >= 0)

    #weekly STEP variables
    @variable(model, Pcharge_STEP[1:Tmax] >= 0)
    @variable(model, Pdecharge_STEP[1:Tmax] >= 0)
    @variable(model, stock_STEP[1:Tmax] >= 0)

#############################

#############################
#Définition de la fonction objectif

    @objective(model, Min, sum(Pnuc .* cnuc) + sum(Pcoal .* ccoal) + sum(Pgaz .* cgaz) + sum(Pfuel .* cfuel) + sum(Phy .* chy) + sum(Pcogen .* ccogen) + (Puns'cuns + Pexc'cexc))

#############################

#############################
#Définition des contraintes

    #balance constraint
    @constraint(model, balance[t in 1:Tmax], sum(Pnuc[t, k] for k in 1:Nnuc) + sum(Pgaz[t, h] for h in 1:Ngaz) + sum(Pcoal[t, g] for g in 1:Ncoal) + sum(Pfuel[t, i] for i in 1:Nfuel) + Pcogen[t] + sum(Phy[t, j] for j in 1:Nhy) + Pres[t] + Pdecharge_STEP[t] + Puns[t] - conso[t] - Pexc[t] - Pcharge_STEP[t] == 0)

    ##
    #constraints for NUCLEAR clusters
    
        #nuc unit Pmax constraints
        @constraint(model, max_nuc[t in 1:Tmax, g in 1:Nnuc], Pnuc[t, g] <= Pmax_nuc[g] * UCnuc[t, g])

        #nuc units Pmin constraints
        @constraint(model, min_nuc[t in 1:Tmax, g in 1:Nnuc], Pmin_nuc[g] * UCnuc[t, g] <= Pnuc[t, g])

        #nuc unit Dmin constraints
        for g in 1:Nnuc
            if (dmin_nuc[g] > 1)
                @constraint(model, [t in 2:Tmax], UCnuc[t, g] - UCnuc[t-1, g] == UPnuc[t, g] - DOnuc[t, g], base_name = "fct_nuc_$g") # 
                @constraint(model, [t in 1:Tmax], UPnuc[t] + DOnuc[t] <= 1, base_name = "UPDOnuc_$g")# peut pas être allumé et éteint en même temps
                @constraint(model, UPnuc[1, g] == 0, base_name = "iniUPnuc_$g") # initialisation
                @constraint(model, DOnuc[1, g] == 0, base_name = "iniDOnuc_$g") # initialisation
                @constraint(model, [t in dmin_nuc[g]:Tmax], UCnuc[t, g] >= sum(UPnuc[i, g] for i in (t-dmin_nuc[g]+1):t), base_name = "dminUPnuc_$g")# temps minimal d'utilisation
                @constraint(model, [t in dmin_nuc[g]:Tmax], UCnuc[t, g] <= 1 - sum(DOnuc[i, g] for i in (t-dmin_nuc[g]+1):t), base_name = "dminDOnuc_$g")
                @constraint(model, [t in 1:dmin_nuc[g]-1], UCnuc[t, g] >= sum(UPnuc[i, g] for i in 1:t), base_name = "dminUPnuc_$(g)_init")# logique, le nombre de fois où il est allumé 
                @constraint(model, [t in 1:dmin_nuc[g]-1], UCnuc[t, g] <= 1 - sum(DOnuc[i, g] for i in 1:t), base_name = "dminDOnuc_$(g)_init")
            end
        end

    ##

    ##
    #constraints for GAS clusters

        #Gaz unit Pmax constraints, on ne peut pas produire plus que Pmax
        @constraint(model, max_gaz[t in 1:Tmax, g in 1:Ngaz], Pgaz[t, g] <= Pmax_gaz[g] * UCgaz[t, g])

        #gaz units Pmin constraints, on ne peut pas l'appeler pour moins que Pmin
        @constraint(model, min_gaz[t in 1:Tmax, g in 1:Ngaz], Pmin_gaz[g] * UCgaz[t, g] <= Pgaz[t, g])

        #gaz unit Dmin constraints
        for g in 1:Ngaz
            if (dmin_gaz[g] > 1)
                @constraint(model, [t in 2:Tmax], UCgaz[t, g] - UCgaz[t-1, g] == UPgaz[t, g] - DOgaz[t, g], base_name = "fct_gaz_$g") # Pour vérifier que UP and DO suivent bien
                @constraint(model, [t in 1:Tmax], UPgaz[t] + DOgaz[t] <= 1, base_name = "UPDOgaz_$g") # peut pas être allumé et éteint en même temps
                @constraint(model, UPgaz[1, g] == 0, base_name = "iniUPgaz_$g") # initialisation A CHANGER pour effet de bord sauf 1ère semaine
                @constraint(model, DOgaz[1, g] == 0, base_name = "iniDOgaz_$g") # initialisation A CHANGER pour effet de bord sauf 1ère semaine
                @constraint(model, [t in dmin_gaz[g]:Tmax], UCgaz[t, g] >= sum(UPgaz[i, g] for i in (t-dmin_gaz[g]+1):t), base_name = "dminUPgaz_$g") #temps minimal d'utilisation dmin_gaz
                @constraint(model, [t in dmin_gaz[g]:Tmax], UCgaz[t, g] <= 1 - sum(DOgaz[i, g] for i in (t-dmin_gaz[g]+1):t), base_name = "dminDOgaz_$g") #temps minimal d'utilisatoin dmin_gaz
                @constraint(model, [t in 1:dmin_gaz[g]-1], UCgaz[t, g] >= sum(UPgaz[i, g] for i in 1:t), base_name = "dminUPgaz_$(g)_init")#logique, le nombre de fois où il est allumé 
                @constraint(model, [t in 1:dmin_gaz[g]-1], UCgaz[t, g] <= 1 - sum(DOgaz[i, g] for i in 1:t), base_name = "dminDOgaz_$(g)_init")
            end
        end

    ##

    ##
    #constraints for COAL clusters

        #Coal unit Pmax constraints
        @constraint(model, max_coal[t in 1:Tmax, g in 1:Ncoal], Pcoal[t, g] <= Pmax_coal[g] * UCcoal[t, g])

        #Coal units Pmin constraints
        @constraint(model, min_coal[t in 1:Tmax, g in 1:Ncoal], Pmin_coal[g] * UCcoal[t, g] <= Pcoal[t, g])

        #Coal unit Dmin constraints
        for g in 1:Ncoal
            if (dmin_coal[g] > 1)
                @constraint(model, [t in 2:Tmax], UCcoal[t, g] - UCcoal[t-1, g] == UPcoal[t, g] - DOcoal[t, g], base_name = "fct_coal_$g") # 
                @constraint(model, [t in 1:Tmax], UPcoal[t] + DOcoal[t] <= 1, base_name = "UPDOcoal_$g") # peut pas être allumé et éteint en même temps
                @constraint(model, UPcoal[1, g] == 0, base_name = "iniUPcoal_$g") # initialisation
                @constraint(model, DOcoal[1, g] == 0, base_name = "iniDOcoal_$g") #initialisatio
                @constraint(model, [t in dmin_coal[g]:Tmax], UCcoal[t, g] >= sum(UPcoal[i, g] for i in (t-dmin_coal[g]+1):t), base_name = "dminUPcoal_$g")#temps minimal d'utilisation
                @constraint(model, [t in dmin_coal[g]:Tmax], UCcoal[t, g] <= 1 - sum(DOcoal[i, g] for i in (t-dmin_coal[g]+1):t), base_name = "dminDOcoal_$g")
                @constraint(model, [t in 1:dmin_coal[g]-1], UCcoal[t, g] >= sum(UPcoal[i, g] for i in 1:t), base_name = "dminUPcoal_$(g)_init")#logique, le nombre de fois où il est allumé 
                @constraint(model, [t in 1:dmin_coal[g]-1], UCcoal[t, g] <= 1 - sum(DOcoal[i, g] for i in 1:t), base_name = "dminDOcoal_$(g)_init")
            end
        end

    ##

    ##
    #constraints for FUEL clusters
        #fuel unit Pmax constraints
        @constraint(model, max_fuel[t in 1:Tmax, g in 1:Nfuel], Pfuel[t, g] <= Pmax_fuel[g] * UCfuel[t, g])

        #fuel units Pmin constraints
        @constraint(model, min_fuel[t in 1:Tmax, g in 1:Nfuel], Pmin_fuel[g] * UCfuel[t, g] <= Pfuel[t, g])

        # fuel unit Dmin constraints
        for g in 1:Nfuel
            if (dmin_fuel[g] > 1)
                @constraint(model, [t in 2:Tmax], UCfuel[t, g] - UCfuel[t-1, g] == UPfuel[t, g] - DOfuel[t, g], base_name = "fct_fuel_$g") # 
                @constraint(model, [t in 1:Tmax], UPfuel[t] + DOfuel[t] <= 1, base_name = "UPDOfuel_$g")#peut pas être allumé et éteint en même temps
                @constraint(model, UPfuel[1, g] == 0, base_name = "iniUPfuel_$g") # initialisation
                @constraint(model, DOfuel[1, g] == 0, base_name = "iniDOfuel_$g") # initialisation
                @constraint(model, [t in dmin_fuel[g]:Tmax], UCfuel[t, g] >= sum(UPfuel[i, g] for i in (t-dmin_fuel[g]+1):t), base_name = "dminUPfuel_$g") # temps minimal d'utilisation
                @constraint(model, [t in dmin_fuel[g]:Tmax], UCfuel[t, g] <= 1 - sum(DOfuel[i, g] for i in (t-dmin_fuel[g]+1):t), base_name = "dminDOfuel_$g")
                @constraint(model, [t in 1:dmin_fuel[g]-1], UCfuel[t, g] >= sum(UPfuel[i, g] for i in 1:t), base_name = "dminUPfuel_$(g)_init") # logique, le nombre de fois où il est allumé 
                @constraint(model, [t in 1:dmin_fuel[g]-1], UCfuel[t, g] <= 1 - sum(DOfuel[i, g] for i in 1:t), base_name = "dminDOfuel_$(g)_init")
            end
        end

    ##

    ##
    #constraints for COGEN clusters (que Pmax constraint)

        @constraint(model, max_cogen[t in 1:Tmax, g in 1:Ncogen], Pcogen[t, g] <= Pmax_cogen[g] * UCcogen[t, g])

    ##

    ##
    #constraints for HYDRO clusters
    
        #on met un pourcentage minimum en dessous duquel on ne doit pas descendre, par heure : e_hy_low
        @constraint(model, pompage_max[t in 1:Tmax, h in 1:Nhy], e_hy[h] >= e_hy_low[t])

        #hydro unit constraints
        @constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t, h] <= Pmax_hy[h])
        #hydro stock constraint
        @constraint(model, stock_hy[h in 1:Nhy], sum(Phy[t, h] for t in 1:Tmax) <= e_hy[h])

    ##

    #weekly STEP ATTENTION ici pour 1 semaine, si élargissement fenêtre (ex: incluer 3 jours avant), peut-être
    @constraint(model, Pcharge_max_STEP[t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
    @constraint(model, Pdecharge_max_STEP[t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
    @constraint(model, init_stock_STEP, stock_STEP[1] == 0)
    @constraint(model, end_Pdecharge_STEP, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
    @constraint(model, Tmax_stock_STEP, stock_STEP[Tmax] == stock_STEP[1])
    @constraint(model, init_Pdecharge_STEP, Pdecharge_STEP[1] == 0)
    @constraint(model, evol_stock_STEP[t in 1:Tmax-1], stock_STEP[t+1] - stock_STEP[t] - rSTEP * Pcharge_STEP[t] + Pdecharge_STEP[t] == 0)
    @constraint(model, stock_max_STEP[t in 1:Tmax], stock_STEP[t] <= 24 * 7 * Pmax_STEP)

#############################

#############################
#Gestion des effets de bord

    # # Gestion des effets de bord pour le nucléaire ATTENTION vérifier que c'est cohérent. RENVOIE UN PROBLÈME
    # if week >= 2
    #     # Semaines suivantes : Charger les états initiaux depuis le fichier effet_bord
    #     for numcentrale in 1:Nnuc
    #         # Déterminer la colonne correspondant à l'unité g
    #         if numcentrale == 1
    #             # Iconuc
    #             UP_col = "A"
    #             DO_col = "B"
    #             UC_col = "C"
    #         elseif numcentrale == 2
    #             # Tabernuc
    #             UP_col = "F"
    #             DO_col = "G"
    #             UC_col = "H"
    #         else
    #             # Necplusultra
    #             UP_col = "K"
    #             DO_col = "I"
    #             UC_col = "J"
    #         end

    #         # Lecture des états UP, DO, et UC depuis le fichier
    #         UPnuc[1, numcentrale] = XLSX.readdata(effet_bord, "Sheet1", "$(UP_col)$(1 + numcentrale)") # UPnuc[1,numcentrale] = au temps 1 pour le nucléaire numcentrale
    #         DOnuc[1, numcentrale] = XLSX.readdata(effet_bord, "Sheet1", "$(DO_col)$(1 + numcentrale)")
    #         UCnuc[1, numcentrale] = XLSX.readdata(effet_bord, "Sheet1", "$(UC_col)$(1 + numcentrale)")
    #     end
    # end

    # # Contraintes de continuité pour les unités nucléaires
    # for g in 1:Nnuc
    #     if dmin_nuc[g] > 1
    #         # Continuité pour les UPnuc et DOnuc
    #         @constraint(model, [t in 2:Tmax],
    #             UCnuc[t, g] - UCnuc[t-1, g] == UPnuc[t, g] - DOnuc[t, g],
    #             base_name = "fct_nuc_$g"
    #         )

    #         # Les UP et DO ne peuvent pas se produire simultanément
    #         @constraint(model, [t in 1:Tmax],
    #             UPnuc[t, g] + DOnuc[t, g] <= 1,
    #             base_name = "UPDOnuc_$g"
    #         )

    #         # Temps minimal d'utilisation après allumage
    #         @constraint(model, [t in dmin_nuc[g]:Tmax],
    #             UCnuc[t, g] >= sum(UPnuc[i, g] for i in (t-dmin_nuc[g]+1):t),
    #             base_name = "dminUPnuc_$g"
    #         )

    #         # Temps minimal d'arrêt après extinction
    #         @constraint(model, [t in dmin_nuc[g]:Tmax],
    #             UCnuc[t, g] <= 1 - sum(DOnuc[i, g] for i in (t-dmin_nuc[g]+1):t),
    #             base_name = "dminDOnuc_$g"
    #         )
    #     end
    # end


    """
    # calcul de Corentin

    # Calcul des temps de Up and Down A tester que pour NUC avant de généraliser aux autres
    for t in 1:Tmax
        for g in 1:Nnuc
            if UPnuc[t, g] == 1
                time = 0
                while UCnuc[t+time, g] == 1
                    time += 1
                    if t + time <= Tmax  # Vérifie que l'index ne dépasse pas les limites
                        timeUPnuc[t+time, g] = time
                    else
                        break
                    end
                end
            end
            if DOnuc[t, g] == 1
                time = 0
                while UCnuc[t+time, g] == 0
                    time += 1
                    if t + time <= Tmax  # Vérifie que l'index ne dépasse pas les limites
                        timeDOnuc[t+time, g] = time
                    else
                        break
                    end
                end
            end
        end

        for g in 1:Ngaz
            if UPgaz[t, g] == 1
                time = 0
                while UCgaz[t+time, g] == 1
                    time += 1
                    if t + time <= Tmax  # Vérifie que l'index ne dépasse pas les limites
                        timeUPgaz[t+time, g] = time
                    else
                        break
                    end
                end
            end
            if DOgaz[t, g] == 1
                time = 0
                while UCgaz[t+time, g] == 0
                    time += 1
                    if t + time <= Tmax  # Vérifie que l'index ne dépasse pas les limites
                        timeDOgaz[t+time, g] = time
                    else
                        break
                    end
                end
            end
        end
    end
    """
#############################

#############################
#Output
#############################

#------------------------------
#print the model
#print(model)
#------------------------------
#solve the model
optimize!(model)
#------------------------------
#Results
#@show termination_status(model)
#@show objective_value(model)

###############


# écrire up, down et depuis combien de temps
#Nuclear
Up_nuc = value.(UPnuc)
Do_nuc = value.(DOnuc)
Uc_nuc = value.(UCnuc)
timeDo_nuc = value.(timeDOnuc)
timeUp_nuc = value.(timeUPnuc)

#Gaz
Up_gaz = value.(UPgaz)
Do_gaz = value.(DOgaz)
Uc_gaz = value.(UCgaz)
timeDo_gaz = value.(timeDOgaz)
timeUp_gaz = value.(timeUPgaz)

#Fuel
Up_fuel = value.(UPfuel)
Do_fuel = value.(DOfuel)
Uc_fuel = value.(UCfuel)
# timeDo_fuel=value.(timeDOfuel)
# timeUp_fuel=value.(timeUPfuel)

#Coal
Up_coal = value.(UPcoal)
Do_coal = value.(DOcoal)
Uc_coal = value.(UCcoal)
# timeDo_coal=value.(timeDOcoal)
# timeUp_coal=value.(timeUPcoal)

#exports results as csv file
nuc_gen = value.(Pnuc)
gaz_gen = value.(Pgaz)
coal_gen = value.(Pcoal)
fuel_gen = value.(Pfuel)
cogen_gen = value.(Pcogen)
hy_gen = value.(Phy)
exc_gen = value.(Pexc)
STEP_charge = -value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)



# Le document de gestion de sortie pour la semaine en cours

    # file handling in write mode
    f = open("Output/results_k_step.csv", "w")
    lines = readlines(f)
    new_lines = map(line -> replace(line, ";" => ","), lines)
    for line in new_lines
        write(f, line * "\n")
    end

    for name in names_nuc
        write(f, "$name,")
    end
    for name in names_gaz
        write(f, "$name,")
    end
    for name in names_coal
        write(f, "$name,")
    end
    for name in names_fuel
        write(f, "$name,")
    end
    write(f, "Cogén,Hydro,STEP turbinage,Puissance résiduelle,STEP pompage,Conso,Conso nette,")

    write(f, "Puissance excès\n")

    for t in 1:Tmax
        for g in 1:Nnuc
            write(f, "$(nuc_gen[t,g]) , ")
        end
        for g in 1:Ngaz
            write(f, "$(gaz_gen[t,g]) , ")
        end
        for g in 1:Ncoal
            write(f, "$(coal_gen[t,g]) , ")
        end
        for g in 1:Nfuel
            write(f, "$(fuel_gen[t,g]) , ")
        end
        for g in 1:Ncogen
            write(f, "$(cogen_gen[t,g]) , ")
        end
        for h in 1:Nhy
            write(f, "$(hy_gen[t,h]) ,")
        end
        write(f, "$(STEP_decharge[t]),$(Pres[t]), $(STEP_charge[t]),$(conso[t]),$(conso[t]-STEP_charge[t]+exc_gen[t]-Pres[t]),")
        write(f, "$(exc_gen[t])\n")

    end

    close(f)
#

# Le document de gestion des bords

    bords = open("Output/results_k_effetbord.csv", "w")

    lines = readlines(bords)
    new_lines = map(line -> replace(line, ";" => ","), lines)
    for line in new_lines
        write(bords, line * "\n")
    end

    # LES TITRES

    for name in names_nuc
        write(bords, "UP" * "$name ,")
        write(bords, "timeUP" * "$name ,")
        write(bords, "DO" * "$name ,")
        write(bords, "timeDO" * "$name ,")
        write(bords, "UC" * "$name ,")
    end

    for name in names_gaz
        write(bords, "UP" * "$name ,")
        write(bords, "timeUP" * "$name ,")
        write(bords, "DO" * "$name ,")
        write(bords, "timeDO" * "$name ,")
        write(bords, "UC" * "$name ,")
    end

    for name in names_coal
        write(bords, "UP" * "$name ,")
        write(bords, "DO" * "$name ,")
        write(bords, "UC" * "$name ,")
    end

    for name in names_fuel
        write(bords, "UP" * "$name ,")
        write(bords, "DO" * "$name ,")
        write(bords, "UC" * "$name ,")
    end

    write(bords, "\n")

    # LES VALEURS

    for t in 1:Tmax
        for g in 1:Nnuc
            write(bords, "$(Up_nuc[t,g]) , ")
            write(bords, "$(timeUp_nuc[t,g]) , ")
            write(bords, "$(Do_nuc[t,g]) , ")
            write(bords, "$(timeDo_nuc[t,g]) , ")
            write(bords, "$(Uc_nuc[t,g]) , ")
        end
        for g in 1:Ngaz
            write(bords, "$(Up_gaz[t,g]) , ")
            write(bords, "$(timeUp_gaz[t,g]) , ")
            write(bords, "$(Do_gaz[t,g]) , ")
            write(bords, "$(timeDo_gaz[t,g]) , ")
            write(bords, "$(Uc_gaz[t,g]) , ")
        end
        for g in 1:Nfuel
            write(bords, "$(Up_fuel[t,g]) , ")
            #write(bords,"$(timeUp_fuel[t,g]) , ")
            write(bords, "$(Do_fuel[t,g]) , ")
            #write(bords,"$(timeDo_fuel[t,g]) , ")
            write(bords, "$(Uc_fuel[t,g]) , ")
        end

        for g in 1:Ncoal
            write(bords, "$(Up_coal[t,g]) , ")
            #write(bords,"$(timeUp_coal[t,g]) , ")
            write(bords, "$(Do_coal[t,g]) , ")
            #write(bords,"$(timeDo_coal[t,g]) , ")
            write(bords, "$(Uc_coal[t,g]) , ")
        end

        write(bords, " \n")

    end


    close(bords)
#
