using IBM.WMQ;
using ServicioMonitor.Helpers;
using System;

namespace ServicioMonitor.Mq
{
    public class TktMq : MqSeries
    {
        new readonly Funcion_Tkt funcion;
        public string strReturn;
        public bool mbFuncionBloque;


        public TktMq()
        {
            funcion = new Funcion_Tkt();
        }
        public double RevisaQueueMq(string strMQCola, MQOPEN lngOpciones)
        {
            //bMQAbrirCola = false;
            double MQRevisaQueue;
            MQRevisaQueue = 0;
            try
            {   //' Se accesa la cola ya sea para leer o escribir

                QUEUE = QMGR.AccessQueue(strMQCola, (int)lngOpciones);

                //bMQAbrirCola = true;
                MQRevisaQueue = QUEUE.CurrentDepth;
                return MQRevisaQueue;
            }
            catch (MQException ex)
            {
                //bMQAbrirCola = false;
                MQRevisaQueue = 0;
                return MQRevisaQueue;
            }
        }

        public bool RecibirMq(string QueueName)
        {
            //cInterfaz escribeLog = new cInterfaz();
            try
            {
                int openOptions = MQC.MQOO_INPUT_AS_Q_DEF + MQC.MQOO_FAIL_IF_QUIESCING;
                QUEUE = QMGR.AccessQueue(QueueName, openOptions);
                MQMessage qMessage = new MQMessage();
                MQGetMessageOptions queueGetMessageOptions = new MQGetMessageOptions();
                queueGetMessageOptions.WaitInterval = 2 * 1000;
                queueGetMessageOptions.Options = MQC.MQGMO_WAIT;
                QUEUE.Get(qMessage, queueGetMessageOptions);

                byte[] byteMessageId = null;
                string strMessageId = "";
                strReturn = "";
                if (qMessage.Format.CompareTo(MQC.MQFMT_STRING) == 0)
                {
                    qMessage.Seek(0);
                    strReturn = System.Text.UTF8Encoding.UTF8.GetString(qMessage.ReadBytes(qMessage.MessageLength));
                    byteMessageId = qMessage.MessageId;
                    strMessageId = qMessage.MessageId.ToString();
                }
                else
                {
                    throw new NotSupportedException(string.Format("Unsupported message format: '{0}' read from queue: {1}.", qMessage.Format, QUEUE));
                }
                string msgRecuperado = "Mensaje recuperado de la queue: " + strReturn;
            }
            catch (MQException MQexp)
            {
                string strCadenaLogMQ = "Error al leer Queue " + MQexp.Reason + " " +
                    MQexp.InnerException + " , " + MQexp.TargetSite + " , " + MQexp.Data +
                    +MQexp.ReasonCode + ", mensaje " + MQexp.Source + "  Error " + MQexp.Message;
                QMGR.Close();
                funcion.Escribe(strCadenaLogMQ, "Mensaje");
                return false;
            }
            finally
            {
                QUEUE.Close();
            }
            return true;
        }

        public bool MQEnviar(string strMQCola, string ls_mensaje)
        {
            try
            {
                MQPutMessageOptions mqsMQOpciones = new MQPutMessageOptions();
                string strMensaje;

                if (AbrirColaMQ(strMQCola, MQOPEN.MQOO_OUTPUT))
                {
                    strMensaje = ls_mensaje;

                    MSG.ClearMessage();
                    MSG.Format = "MQSTR   ";
                    MSG.MessageType = MQC.MQMT_DATAGRAM;
                    MSG.WriteLong(strMensaje.Trim().Length);
                    MSG.WriteUTF(strMensaje);
                    QUEUE.Put(MSG, mqsMQOpciones);

                    CerrarColaMQ();
                }
            }
            catch (Exception)
            {
                return false;
                throw;
            }
            return true;
        }
    }
}
