using IBM.WMQ;
using ServicioMonitor.Helpers;
using System;

namespace ServicioMonitor.Mq
{
    public class MqSeries
    {
        public bool blnConectado;

        public enum MQOPEN
        {
            MQOO_INPUT_AS_Q_DEF = 0x1,
            MQOO_INPUT_SHARED = 0x2,
            MQOO_INPUT_EXCLUSIVE = 0x4,
            MQOO_BROWSE = 0x8,
            MQOO_OUTPUT = 0x10,
            MQOO_INQUIRE = 0x20,
            MQOO_SET = 0x40,
            MQOO_BIND_ON_OPEN = 0x4000,
            MQOO_BIND_NOT_FIXED = 0x8000,
            MQOO_BIND_AS_Q_DEF = 0x0,
            MQOO_SAVE_ALL_CONTEXT = 0x80,
            MQOO_PASS_IDENTITY_CONTEXT = 0x100,
            MQOO_PASS_ALL_CONTEXT = 0x200,
            MQOO_SET_IDENTITY_CONTEXT = 0x400,
            MQOO_SET_ALL_CONTEXT = 0x800,
            MQOO_ALTERNATE_USER_AUTHORITY = 0x1000,
            MQOO_FAIL_IF_QUIESCING = 0x2000
        };

        public enum MQE
        {
            MQTRANSPORT_PROPERTY = '1',
            MQHOST_NAME_PROPERTY = '0',
            MQPORT_PROPERTY = 1414,
            MQCHANNEL_PROPERTY = '0',
        };


        public MQQueueManager QMGR = null; //mqsManager
        public MQQueueManager QMGR1 = null;
        public MQQueue QUEUE = null; //MQQMonitorLectura
        public MQQueue QUEUE1 = null;
        public MQPutMessageOptions pmo = null;
        public MQMessage MSG = null;

        public Funcion funcion;


        public MqSeries()
        {
            funcion = new Funcion();
        }

        /// <summary>
        /// Cerrar cola MqSeries
        /// </summary>
        /// <returns></returns>
        public bool CerrarColaMQ()
        {
            try
            {
                if (QUEUE != null)
                {
                    if (QUEUE.IsOpen)
                    {
                        QUEUE.Close();
                        return true;
                    }
                }
                return false;
            }
            catch (MQException ex)
            {
                funcion.Escribe(ex, "MQException");
                return false;
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex);
                return false;
            }
        }

        /// <summary>
        /// Desconectar el MqSeries
        /// </summary>
        /// <returns></returns>
        public bool DesconectarMQ()
        {
            bool DesconectarMQ;
            try
            {
                if (blnConectado)
                {
                    QMGR.Disconnect();
                    DesconectarMQ = true;
                }
                else
                {
                    DesconectarMQ = false;
                }
            }
            catch (MQException mq_ex)
            {
                DesconectarMQ = false;
                funcion.Escribe(mq_ex);
            }
            catch (Exception ex)
            {
                DesconectarMQ = false;
                funcion.Escribe(ex);
            }
            return DesconectarMQ;
        }

        /// <summary>
        /// Conectar al MqSeries
        /// </summary>
        /// <param name="strQueueManagerName"></param>
        /// <param name="queueManager"></param>
        /// <returns></returns>
        public bool ConectarMQ(string strQueueManagerName)
        {
            try
            {
                QMGR = new MQQueueManager(strQueueManagerName);

                return true;
            }
            catch (MQException mq_ex)
            {
                funcion.Escribe(mq_ex);
                return false;

            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
                return false;

            }
        }

        /// <summary>
        /// Enviar mensaje a la cola MQ
        /// </summary>
        /// <returns></returns>
        public bool EnviarMensajeMQ(string strMQCola)
        {
            long lngMqOpen;
            bool EnviarMensajeMQ = false;
            try
            {
                lngMqOpen = (long)MQOPEN.MQOO_OUTPUT;

                if (AbrirColaMQ(strMQCola, (MQOPEN)lngMqOpen))
                {
                    pmo = new MQPutMessageOptions();
                    pmo.Options = MQC.MQPMO_SYNCPOINT;
                    MSG = new MQMessage();
                    QUEUE.Put(MSG, pmo);
                    EnviarMensajeMQ = true;
                }


            }
            catch (MQException mq_ex)
            {
                EnviarMensajeMQ = false;
                funcion.Escribe(mq_ex, "Error");
            }
            catch (Exception ex)
            {
                EnviarMensajeMQ = false;
                funcion.Escribe(ex, "Error");
            }


            return EnviarMensajeMQ;
        }

        /// <summary>
        /// Abrir cola del MQ
        /// </summary>
        /// <returns></returns>
        public bool AbrirColaMQ(string strMQCola, MQOPEN lngOpciones)
        {
            bool AbriColaMQ;
            try
            {
                QUEUE = QMGR.AccessQueue(strMQCola,
                   (int)lngOpciones
                );

                AbriColaMQ = true;
            }
            catch (MQException mq_ex)
            {
                AbriColaMQ = false;
                funcion.Escribe(mq_ex, "Error");
            }
            catch (Exception ex)
            {
                AbriColaMQ = false;
                funcion.Escribe(ex, "Error");
            }

            return AbriColaMQ;

        }




    }
}
