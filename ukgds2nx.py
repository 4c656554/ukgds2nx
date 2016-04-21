#  Copyright (c) 2015 Lee J. Thomas <ThomasL62@cf.ac.uk>
#
#  Permission is hereby granted, free of charge, to any person obtaining a copy
#  of this software and associated documentation files (the "Software"), to
#  deal in the Software without restriction, including without limitation the
#  rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
#  sell copies of the Software, and to permit persons to whom the Software is
#  furnished to do so, subject to the following conditions:
#
#  The above copyright notice and this permission notice shall be included in
#  all copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
#  IN THE SOFTWARE.
#------------------------------------------------------------------------------

""" Defines a parser for extracting network data stored according to the United
    Kingdom Generic Distribution System (UKGDS) format in an Excel spreadsheet.

    Converts the data to NetworkX MultiGraph format and saves as a .gpickle file.

    @see: Foote C., Djapic P,. Ault G., Mutale J., Burt G., Strbac G., "United
    Kingdom Generic Distribution System (UKGDS) Software Tools", The Centre for
    Distributed Generation and Sustainable Electrical Energy, February 2006
"""

#------------------------------------------------------------------------------
#  Imports:
#------------------------------------------------------------------------------

# Operating system routines.
import os

# Numpy lib for convenient data storage/access
import numpy as np

# Pandas lib for convenient data storage/access
import pandas as pd

# Matplotlib lib for plotting
import matplotlib.pyplot as plt
from matplotlib import colors
from matplotlib.mlab import griddata

import networkx as nx


#------------------------------------------------------------------------------
#  "ukgds2nx" function:
#------------------------------------------------------------------------------

def ukgds2nx(infile, outfile):
    """ Extracts UKGDS data from an Excel spreadsheet and writes to gpickle file
    """

    # Check that the path to the inputfile is valid.
    if not os.path.isfile(infile):
        raise NameError, "%s is not a valid filename" % infile
    
    # System sheet ---------------------------------------------------------------
    try:
        sys_data=pd.read_excel(infile,'System',skiprows=28)
    except:
        print 'No system tab'
    # Bus sheet ---------------------------------------------------------------
    try:
        bus_data=pd.read_excel(infile,'Buses',skiprows=28,parse_cols=range(1,12))
    except:
        print 'No Buses tab'
    # Loads sheet ---------------------------------------------------------------
    try:
        load_data=pd.read_excel(infile,'Loads',skiprows=28,parse_cols=range(1,7))
    except:
        print 'No Loads tab'
    # Generators sheet ---------------------------------------------------------------
    try:
        gen_data=pd.read_excel(infile,'Generators',skiprows=28,parse_cols=range(1,17))
    except:
        print 'No Generators tab'
    # Transformers sheet ---------------------------------------------------------------
    try:
        tx_data=pd.read_excel(infile,'Transformers',skiprows=28,parse_cols=range(1,23))
    except:
        print 'No Transformers tab'
    # IndGenerators sheet ---------------------------------------------------------------
    try:
        indgen_data=pd.read_excel(infile,'IndGenerators',skiprows=28,parse_cols=range(1,23))
    except:
        print 'No IndGenerators tab'
    # Shunts sheet ---------------------------------------------------------------
    try:
        shunts_data=pd.read_excel(infile,'Shunts',skiprows=28,parse_cols=range(1,6))
    except:
        print 'No Shunts tab'
    # Branches sheet ---------------------------------------------------------------
    try:
        branch_data=pd.read_excel(infile,'Branches',skiprows=28,parse_cols=range(1,15))
    except:
        print 'No Branches tab'
    # Taps sheet ---------------------------------------------------------------
    try:
        taps_data=pd.read_excel(infile,'Taps',skiprows=28,parse_cols=range(1,23))
    except:
        print 'No Taps tab'
    
    
    #--------------------------------------------------------------------------
    # Create networkx graph 
    #--------------------------------------------------------------------------

    G = nx.MultiGraph()
    dupedgeind=0
    ind=0

    for index,bus in bus_data.iterrows():
        G.add_node(int(bus['BNU']),attr_dict=dict(bus))
        ind+=1
    for index,b in branch_data.iterrows():
        if G.has_edge(int(b['CFB']),int(b['CTB'])):
            dupedgeind=dupedgeind+1
            G.add_edge(int(b['CFB']),int(b['CTB']),key=dupedgeind,len=b['CLE'],attr_dict=dict(b))
        else:
            G.add_edge(int(b['CFB']),int(b['CTB']),key=dupedgeind,len=b['CLE'],attr_dict=dict(b))
            dupedgeind=0
    dupedgeind=0
    for index,t in tx_data.iterrows():
        if G.has_edge(int(t['TFB']),int(t['TTB'])):
            dupedgeind=dupedgeind+1
            G.add_edge(int(t['TFB']),int(t['TTB']),key=dupedgeind,len=0.01,attr_dict=dict(t))
        else:
            G.add_edge(int(t['TFB']),int(t['TTB']),key=dupedgeind,len=0.01,attr_dict=dict(t))
            dupedgeind=0

    ind=0

    for index,load in load_data.iterrows():
        G.add_node(int(load['LBN']),attr_dict=dict(load))
        ind+=1
    ind=0
    for index,gen in gen_data.iterrows():
        G.add_node(int(gen['GBN']),attr_dict=dict(gen))
        ind+=1
    
    G.add_node(int(sys_data.Value[sys_data.Symbol=='SSB'].values[0]),isSlack=True,BaseMVA=sys_data.Value[sys_data.Symbol=='SMB'].values[0],desc=sys_data.Value[sys_data.Symbol=='STD'].values[0])
    ind=0


    #--------------------------------------------------------------------------
    #  Write the text to the gpickle file:
    #--------------------------------------------------------------------------
    nx.write_gpickle(G,outfile+'.gpickle')




def plotmap(G,fileref='ukgds2nxFigure',graphprogram='neato',tag=''):
    """ Generates node positions and plots UKGDS NetworkX Graph. No indication for parallel edges (lines) between nodes. Saves as a .png
    """
    voltages = np.unique([G.node[x]['BBV'] for x,y in G.edges()])
    colours = ['b','b','k','#808000','g','r','r','r','y']
    thickness = [0.5*x for x in [1,1,1,1,1,1,1,1,1]]
    alpha = [0.5*x for x in [1,1,1,1,1,1,1,1,1]]
    colourmaps = [plt.cm.Blues,plt.cm.Blues,plt.cm.Greys,plt.cm.OrRd,plt.cm.summer,plt.cm.autumn,plt.cm.autumn,plt.cm.autumn,plt.cm.autumn]
    pos=nx.graphviz_layout(G,prog=graphprogram) #This generates the node positions
    for p,q in pos.iteritems():
        G.add_node(p,pos=q)

        
    notxedges = [(x,y,d) for x,y,d in G.edges(data=True) if 'CLE' in d]
    labels=dict(((x,y),str(d['CLE'])+'km') for x,y,d in G.edges(data=True) if 'CLE' in d) 
    
    #--------------------------------------------------------------------------
    #  Plot for 11kV nodes
    #--------------------------------------------------------------------------
    nodes = [x for x,d in G.nodes(data=True) if d['BBV']==11]
    #print nodes
    SG = G.subgraph(nodes)
    poslabels={}
    for sgn in SG.nodes():
        x,y = pos[sgn]
        poslabels[sgn] = (x,y+100)
    nodelabels = dict((x,x) for x,y in SG.nodes(data=True))
    #print nx.get_node_attributes(SG,'pos')
    #pos=[dict(n,pp) for n,pp in zip(SG.nodes(),[(p[0],p[1]+1) for x,p in nx.get_node_attributes(SG,'pos')])]
    ##print pos
    #nx.draw_networkx_labels(SG,pos=poslabels,labels=nodelabels,font_size=8,alpha=0.6) 
    nx.draw_networkx_edges(SG,nx.get_node_attributes(SG,'pos'),edgelist=SG.edges(),edge_color='r',style='-',width=3,with_labels=False,alpha=0.4)
    nx.draw(SG,nx.get_node_attributes(SG,'pos'),node_size=10,with_labels=False,node_color='r',alpha=0.4)
    
    #--------------------------------------------------------------------------
    #  Plot for 33kV nodes
    #--------------------------------------------------------------------------
    nodes = [x for x,d in G.nodes(data=True) if d['BBV']==33]
    #print nodes
    SG = G.subgraph(nodes)
    poslabels={}
    for sgn in SG.nodes():
        x,y = pos[sgn]
        poslabels[sgn] = (x,y-100)
    nodelabels = dict((x,x) for x,y in SG.nodes(data=True))
    #print nx.get_node_attributes(SG,'pos')
    nx.draw_networkx_labels(SG,pos=poslabels,labels=nodelabels,font_size=8,alpha=0.8) 
    nx.draw(SG,nx.get_node_attributes(SG,'pos'),node_size=20,with_labels=False,node_color='g',edge_color='g',style='-',width=3,alpha=0.4)
    
    #--------------------------------------------------------------------------
    #  Plot for 132kV nodes
    #--------------------------------------------------------------------------
    nodes = [x for x,d in G.nodes(data=True) if d['BBV']==132]
    #print nodes
    SG = G.subgraph(nodes)
    poslabels={}
    for sgn in SG.nodes():
        x,y = pos[sgn]
        poslabels[sgn] = (x,y+100)
    nodelabels = dict((x,x) for x,y in SG.nodes(data=True))
    #print nx.get_node_attributes(SG,'pos')
    nx.draw_networkx_labels(SG,pos=poslabels,labels=nodelabels,font_size=10,alpha=0.8) 
    nx.draw_networkx_edges(SG,nx.get_node_attributes(SG,'pos'),edgelist=SG.edges(),edge_color='k',style='-',width=3,with_labels=False,alpha=0.4)
    nx.draw(SG,nx.get_node_attributes(SG,'pos'),node_size=30,with_labels=False,node_color='k',alpha=0.4)
    #nx.draw_networkx_edge_labels(G,nx.get_node_attributes(G,'pos'),edge_labels=labels,rotate=False,font_size=7)    
   
    plt.savefig(fileref+str(graphprogram)+str(tag)+'.png',dpi=300))
    
   
