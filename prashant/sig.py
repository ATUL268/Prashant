from pprint import pprint
import pdb 
import pandas as pd
import time
import datetime
import xlwings as xw
import zrd_login

kite = zrd_login.kite



# Full Updated by EOD Moday 3:30-----By Atul


def LTP(name):
	last_price = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
	return last_price


def LTP_NFO(name):
	last_price = kite.ltp(['NFO:'+ name])['NFO:'+ name]['last_price']
	return last_price


#-------------------------------------------------Code--------------------------code----------------------------




wb = xw.Book('signals.xlsx')
sht = wb.sheets['Sheet1']

wb1 = xw.Book('Trade.xlsx')
trd = wb1.sheets['Sheet1']
rec = wb1.sheets['Sheet2']

step_value  = {'NIFTY 50' : 50 , 'NIFTY BANK' : 100}
multiplier = 0
sht.range("b2:q100").value = None
trd.range("b2:za100").value = None
 
expiry = '23615'
trigger  = 1.01 # It can be changed a per user and 1 is added to add percentage
sl = .10
sl_t2 = 0.02
target = .01
trail = 0.01
tradeno = 0
final = {}

Quantity = {'NIFTY 50': {'qty': '50'}, 'NIFTY BANK': {'qty': '25'}}

mohar = {'NIFTY 50': {'my_name': 'NIFTY'}, 'NIFTY BANK': {'my_name': 'BANKNIFTY'}}


temp = {'ATM_STRIKE' : None,'ATM_CE_LTP' : None , 'ATM_PE_LTP': None , 'ATM_CE_TGV' : None , 'ATM_PE_TGV' : None ,  'Picked' : None ,"Time" : None , 'QTY': None ,'Buy_Price_CE' :None ,'Traded_CE': None ,'Stop_Loss_CE': None,  'Target_CE': None, 'New_Buy_Price_CE' : None , 'Next_Trail_Price_CE': None ,'Trade_Completed_CE' : None ,'Sell_Price_CE' : None ,'PNL_CE' : None ,'Remark_CE' : None ,  'CE_Strike' : None, 'Buy_Price_PE' :None ,'Traded_PE': None ,'Stop_Loss_PE': None,  'Target_PE': None, 'New_Buy_Price_PE' : None , 'Next_Trail_Price_PE' : None ,'Sell_Price_PE' : None ,   'PNL_PE' : None ,'Remark_PE' : None , 'Trade_Completed_PE' : None   ,   'Trade_Completed_PE' : None ,  'PE_Strike' : None } 

status = {}


watchlist = sht.range(f"a{2}:a{3}").value

for name in watchlist:
		status[name] = temp.copy()


# print('Before While')
print(pd.DataFrame(status).T)

trd.range('A1').value  = pd.DataFrame(status)

# pdb.set_trace()
print('While loop Start Good !!!!!...................................')

while True:
	# print('Me While Me Hu......................................')
	try:
		time.sleep(0.25)
		ctime = datetime.datetime.now().time()
		
		watchlist = sht.range(f"a{2}:a{3}").value
		get_signal = sht.range(f"b{2}:b{3}").value
		
		idx = 0
		for name  in watchlist:
			# print(name)
			# pdb.set_trace()
			# pdb.set_trace()
			
			 
			signal = get_signal[idx]

			if signal is not None and sht.range(f"C{idx + 2}").value is None and status[name]['Picked'] is None:
				# print(name)
				# print('HELLLLLLLLLLLLLL')
				sht.range(f"C{idx + 2}").value = 'C' + str(idx + 2)
				

				ltp  = LTP(name=name)

				atm_strike = round(ltp/step_value[name])* step_value[name] + multiplier*step_value[name]
				print(atm_strike)
				time.sleep(0.1)
			

				# pdb.set_trace()
				ATM_CE =  mohar[name]['my_name'] + expiry  + str(atm_strike) + 'CE'
				ATM_PE =  mohar[name]['my_name'] + expiry  + str(atm_strike) + 'PE'

				print(ATM_CE , ATM_PE)

				# if name == 'NIFTY 50':

				# 	ATM_CE =  'NIFTY' + expiry  + str(atm_strike) + 'CE'
				# 	ATM_PE =  'NIFTY' + expiry  + str(atm_strike) + 'PE'
				# pdb.set_trace()
				status[name]['ATM_STRIKE'] = atm_strike
				status[name]['CE_Strike'] = ATM_CE
				status[name]['PE_Strike'] = ATM_PE
				status[name]['ATM_CE_LTP'] = LTP_NFO(ATM_CE)
				status[name]['ATM_PE_LTP'] = LTP_NFO(ATM_PE)

				status[name]['ATM_CE_TGV'] = round(status[name]['ATM_CE_LTP'] * trigger , 2) # TGV = Trigger value CE
				status[name]['ATM_PE_TGV'] = round(status[name]['ATM_PE_LTP'] * trigger , 2) # TGV = Trigger value PE

				status[name]['Picked'] = 'YES'
				status[name]['Time'] = str(ctime)
				print(pd.DataFrame(status).T)
				# pdb.set_trace()
				trd.range('A1').value  = pd.DataFrame(status)
				# print("First.......................IF.................1..........................")
				# print("First.......................IF.................1..........................")
				# print("First.......................IF.................1..........................")
				# pdb.set_trace()


			if status[name]['Picked'] == 'YES': # it will run only when signal came into pitcher
				# ltp_CE  = LTP_NFO(name=ATM_CE)
				# ltp_PE  = LTP_NFO(name=ATM_PE)
				status[name]['ATM_CE_LTP'] = LTP_NFO(mohar[name]['my_name'] + expiry  + str(status[name]['ATM_STRIKE']) + 'CE')
				status[name]['ATM_PE_LTP'] = LTP_NFO(mohar[name]['my_name'] + expiry  + str(status[name]['ATM_STRIKE']) + 'PE')
				# print(ltp_CE)
				# print(ltp_PE)
				# print(pd.DataFrame(status).T)
				trd.range('A1').value  = pd.DataFrame(status)
				# print("Second.......................IF.................2..........................")
				# print("Second.......................IF.................2..........................")
				# print("Second.......................IF.................2..........................")

				# pdb.set_trace()

				if status[name]['ATM_CE_LTP'] > status[name]['ATM_CE_TGV'] and status[name]['Traded_CE'] is None:
					orderid_CE = kite.place_order(variety=kite.VARIETY_REGULAR, exchange=kite.EXCHANGE_NFO, tradingsymbol= status[name]['CE_Strike'], transaction_type=kite.TRANSACTION_TYPE_BUY, quantity= Quantity[name][qty], product=kite.PRODUCT_MIS, order_type=kite.ORDER_TYPE_MARKET, price=None, validity=None, disclosed_quantity=None, trigger_price=None, squareoff=None, stoploss=None, trailing_stoploss=None, tag=None)

					# place order 
					# status[name]['Buy_Price_CE'] = status[name]['ATM_CE_LTP'] 
					status[name]['QTY'] = Quantity[name][qty]
					status[name]['Traded_CE'] = 'YES'
					status[name]['Stop_Loss_CE'] = (status[name]['Buy_Price_CE']) - (status[name]['Buy_Price_CE'] * sl)
					status[name]['Target_CE'] = (status[name]['Buy_Price_CE']) + (status[name]['Buy_Price_CE'] * (target))
					status[name]['New_Buy_Price_CE'] = status[name]['Buy_Price_CE']
					trd.range('A1').value  = pd.DataFrame(status)
					print("Second.......................IF.................2A..........................")
					order_history_CE = kite.order_history(Status['orderid_CE'])


					try:
						x = 0
						for order in order_history_CE:
							if order_history_CE[x]['status'] == 'COMPLETE':
								status[name]['Buy_Price_CE'] = (order_history_CE[x]['average_price'])
							x = x + 1

					except Exception as e:
						print(e)
						continue
					

				if status[name]['ATM_PE_LTP'] > status[name]['ATM_PE_TGV'] and status[name]['Traded_PE'] is None:
					orderid_PE = kite.place_order(variety=kite.VARIETY_REGULAR, exchange=kite.EXCHANGE_NFO, tradingsymbol=status[name]['PE_Strike'], transaction_type=kite.TRANSACTION_TYPE_BUY, quantity= Quantity[name][qty], product=kite.PRODUCT_MIS, order_type=kite.ORDER_TYPE_MARKET, price=None, validity=None, disclosed_quantity=None, trigger_price=None, squareoff=None, stoploss=None, trailing_stoploss=None, tag=None)

					#place order 
					# status[name]['Buy_Price_PE'] = status[name]['ATM_PE_LTP']
					status[name]['QTY'] = Quantity[name][qty]
					status[name]['Traded_PE'] = 'YES'
					status[name]['Stop_Loss_PE'] = (status[name]['Buy_Price_PE']) - (status[name]['Buy_Price_PE'] * sl)
					status[name]['Target_PE'] = (status[name]['Buy_Price_PE']) + (status[name]['Buy_Price_PE'] * (target))
					status[name]['New_Buy_Price_PE'] = status[name]['Buy_Price_PE']
					trd.range('A1').value  = pd.DataFrame(status)
					print("Second.......................IF.................2B..........................")
					order_history_PE = kite.order_history(Status['orderid_PE'])



					try:
						x = 0
						for order in order_history_PE:
							if order_history_PE[x]['status'] == 'COMPLETE':
								status[name]['Buy_Price_PE'] = (order_history_PE[x]['average_price'])
							x = x + 1

					except Exception as e:
						print(e)
						continue
					



			if status[name]['Traded_CE'] == 'YES' and status[name]['Picked'] == 'YES' and status[name]['Trade_Completed_CE'] is None:

				x = (status[name]['New_Buy_Price_CE']) + (trail)*(status[name]['New_Buy_Price_CE'])
				
				status[name]['Next_Trail_Price_CE'] = x
				trd.range('A1').value  = pd.DataFrame(status)


				if status[name]['ATM_CE_LTP'] > x : # Stop loss Trailing
					status[name]['New_Buy_Price_CE'] = x
					status[name]['Stop_Loss_CE']  = round(status[name]['Stop_Loss_CE'] + (((status[name]['New_Buy_Price_CE'])*(sl_t2))/2),2)
					trd.range('A1').value  = pd.DataFrame(status)
					# pdb.set_trace()


				if ((status[name]['ATM_CE_LTP'] < status[name]['Stop_Loss_CE']) or (status[name]['ATM_CE_LTP'] > status[name]['Target_CE'])) and (status[name]['PNL_CE'] is None):

					sell_orderid_CE_2 = kite.place_order(variety=kite.VARIETY_REGULAR, exchange=kite.EXCHANGE_NFO, tradingsymbol= status[name]['CE_Strike'], transaction_type=kite.TRANSACTION_TYPE_SELL, quantity= status[name]['QTY'] , product=kite.PRODUCT_MIS, order_type=kite.ORDER_TYPE_MARKET, price=None, validity=None, disclosed_quantity=None, trigger_price=None, squareoff=None, stoploss=None, trailing_stoploss=None, tag=None)

					status[name]['Sell_Price_CE'] = status[name]['ATM_CE_LTP']
					status[name]['PNL_CE'] = (status[name]['Sell_Price_CE'] - status[name]['Buy_Price_CE']) * status[name]['QTY']
					status[name]['Trade_Completed_CE'] = 'YES'
					trd.range('A1').value  = pd.DataFrame(status)

					# pdb.set_trace()
					


					if (status[name]['Sell_Price_CE'] < status[name]['Stop_Loss_CE']):

						status[name]['Remark_CE'] = 'StopLoss_Hit'
						trd.range('A1').value  = pd.DataFrame(status)
						final[tradeno] = status[name]
						rec.range('A1').value  = pd.DataFrame(final).T
						tradeno = tradeno + 1





					if (status[name]['Sell_Price_CE'] > status[name]['Target_CE']):

						status[name]['Remark_CE'] = 'Target_Hit'
						trd.range('A1').value  = pd.DataFrame(status)
						final[tradeno] = status[name]
						rec.range('A1').value  = pd.DataFrame(final).T
						tradeno = tradeno + 1


			if status[name]['Traded_PE'] == 'YES' and status[name]['Picked'] == 'YES' and status[name]['Trade_Completed_PE'] is None:

				x = (status[name]['New_Buy_Price_PE']) + (trail)*(status[name]['New_Buy_Price_PE'])
				
				status[name]['Next_Trail_Price_PE'] = x
				trd.range('A1').value  = pd.DataFrame(status)


				if status[name]['ATM_PE_LTP'] > x :
					status[name]['New_Buy_Price_PE'] = x
					status[name]['Stop_Loss_PE']  = round(status[name]['Stop_Loss_PE'] + (((status[name]['New_Buy_Price_PE'])*(sl_t2))/2),2)

					trd.range('A1').value  = pd.DataFrame(status)
					# pdb.set_trace()


				if ((status[name]['ATM_PE_LTP'] < status[name]['Stop_Loss_PE']) or (status[name]['ATM_PE_LTP'] > status[name]['Target_PE'])) and (status[name]['PNL_PE'] is None):

					sell_orderid_CE = kite.place_order(variety=kite.VARIETY_REGULAR, exchange=kite.EXCHANGE_NFO, tradingsymbol= status[name]['PE_Strike'], transaction_type=kite.TRANSACTION_TYPE_SELL, quantity= status[name]['QTY'] , product=kite.PRODUCT_MIS, order_type=kite.ORDER_TYPE_MARKET, price=None, validity=None, disclosed_quantity=None, trigger_price=None, squareoff=None, stoploss=None, trailing_stoploss=None, tag=None)

					status[name]['Sell_Price_PE'] = status[name]['ATM_PE_LTP']
					status[name]['PNL_PE'] = (status[name]['Sell_Price_PE'] - status[name]['Buy_Price_PE']) * status[name]['QTY']
					status[name]['Trade_Completed_PE'] = 'YES'
					trd.range('A1').value  = pd.DataFrame(status)
					# pdb.set_trace()

					


					if (status[name]['Sell_Price_PE'] < status[name]['Stop_Loss_PE']):

						status[name]['Remark_PE'] = 'StopLoss_Hit'
						trd.range('A1').value  = pd.DataFrame(status)
						final[tradeno] = status[name]
						rec.range('A1').value  = pd.DataFrame(final).T
						tradeno = tradeno + 1



					if (status[name]['Sell_Price_PE'] > status[name]['Target_PE']):

						status[name]['Remark_PE'] = 'Target_Hit'
						trd.range('A1').value  = pd.DataFrame(status)
						final[tradeno] = status[name]
						rec.range('A1').value  = pd.DataFrame(final).T
						tradeno = tradeno + 1

			if status[name]['Remark_CE'] == 'Target_Hit'  or status[name]['Remark_PE'] =='Target_Hit' :
				trd.range('A1').value  = pd.DataFrame(status)
				status[name] = temp.copy()


			if status[name]['Remark_CE'] == 'StopLoss_Hit'  and status[name]['Remark_PE'] =='StopLoss_Hit' :
				trd.range('A1').value  = pd.DataFrame(status)
				status[name] = temp.copy()

			# pdb.set_trace()		
				
			idx += 1 
			# print(idx)

	except Exception as e:
		print(e)

		# pdb.set_trace()




		

	


